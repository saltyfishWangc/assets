package com.example.msyd.assets;

import com.alibaba.fastjson.JSONObject;
import com.example.msyd.assets.config.Config;
import com.example.msyd.assets.formatter.ColumnFormatter;
import com.example.msyd.assets.formatter.DateTimeColumnFormatter;
import com.example.msyd.assets.util.AsposeUtil;
import com.example.msyd.assets.util.TimeMonitor;
import com.msjr.dfs.client.DFSClient;
import com.msjr.dfs.commons.DFSMessage;
import lombok.extern.log4j.Log4j;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.CountDownLatch;
import java.util.logging.Level;
import java.util.regex.Pattern;

/**
 * @author
 * @Description:
 * @date 2022/6/17 15:21
 */
@Component
@Slf4j
public class Creator {

    @Autowired
    Config config;

    // 身份证号-姓名
    private static Map<String, String> idNo2Name = new HashMap<>();

    // 身份证号-用户编号
    private static Map<String, String> idNo2CustNo = new HashMap<>();

    // sql总条数
    private static int count = 0;

    // 遇到数据量大的时候，dba那边自己会将文件拆成4000条脚本一个文件，为减少dba操作，这里控制下每个文件的sql条数，生成多个文件
    private static int maxCountPerFile = 20000;

    // 是否需要在每个文件的头尾加上自动提交
    private static boolean isAppendSqlFix = false;

    // 最终生成的数据变更文件存放地址 后缀
    private static String getDataModSaveFilePathSuffix = ".sql";

    // 用户初始化脚本模板
    private static String 放款凭证sqlTemplate = "INSERT INTO `cmis`.`assets_loan_cont_info`(`trade_no`, `loan_no`, `loan_name`, `id_no`, `apple_time`, `name`, `bank_name`, `account`, `upper_num`, `lower_num`, `success_time`, `receipt_no`, `file_id`) VALUES ('交易号', '借据编号', '借款人姓名', '证件号码', '申请时间', '收款方全称', '收款方开户单位', '收款方账号', '大写金额', '金额小写', '成功时间', '电子回单编号', '');";
    private static String 扣款凭证sqlTemplate = "INSERT INTO `cmis`.`assets_deduct_cont_info`(`trade_no`, `loan_no`, `bank_info`, `amount`, `fee`, `trade_sts`, `create_time`, `success_time`, `file_id`) VALUES ('交易号', '借据编号', '扣款账户信息', '扣款金额', '手续费', '交易状态', '创建时间', '成功时间', '');";

    private static String 放款凭证delSqlTemplate = "delete from `cmis`.`assets_loan_cont_info` where loan_no = '借据编号';";


    // 文件列表按照文件名根据以下排序
//    private static final String[] fileNameSortedKey = {"放款凭证", "扣款凭证"};
    private static final String[] fileNameSortedKey = {"扣款凭证"};

    private static Map<String, String> sqlTemplateMapping = new HashMap<>();

    private static String currentFileName = "";

    /**
     * 列格式化
     *  用法：根据文件名映射对应的列格式化容器，容器里面根据列名映射对应的列格式化类
     */
    private static Map<String, Map<String, ColumnFormatter>> rowFormatMapping = new HashMap<>();

    static {
//        sqlTemplateMapping.put("放款凭证", "INSERT INTO `cmis`.`assets_loan_cont_info`(`trade_no`, `loan_no`, `loan_name`, `id_no`, `apple_time`, `name`, `bank_name`, `account`, `upper_num`, `lower_num`, `success_time`, `receipt_no`, `file_id`) VALUES ('交易号', '借据编号', '借款人姓名', '证件号码', '申请时间', '收款方全称', '收款方开户单位', '收款方账号', '大写金额', '金额小写', '成功时间', '交易号', '');");
        sqlTemplateMapping.put("放款凭证", "update `cmis`.`assets_loan_cont_info` set file_id = '文件ID' where trade_no = '交易号';");
//        sqlTemplateMapping.put("扣款凭证", "INSERT INTO `cmis`.`assets_deduct_cont_info`(`trade_no`, `loan_no`, `bank_info`, `amount`, `fee`, `trade_sts`, `create_time`, `success_time`, `file_id`) VALUES ('交易号', '借据编号', '扣款账户信息', '扣款金额', '手续费', '交易状态', '创建时间', '成功时间', '');");
        sqlTemplateMapping.put("扣款凭证", "update `cmis`.`assets_deduct_cont_info` set file_id = '文件ID' where trade_no = '交易号';");

        // 根据文件名初始化对应的列格式化类
        for (String fileName : fileNameSortedKey) {
            Map<String, ColumnFormatter> columnFormatterMap = new HashMap<>();
            if ("放款凭证".equals(fileName)) {
                columnFormatterMap.put("申请时间", new DateTimeColumnFormatter());
                columnFormatterMap.put("成功时间", new DateTimeColumnFormatter());
            } else if ("扣款凭证".equals(fileName)) {
                columnFormatterMap.put("创建时间", new DateTimeColumnFormatter());
                columnFormatterMap.put("成功时间", new DateTimeColumnFormatter());
            }
            rowFormatMapping.put(fileName, columnFormatterMap);
        }

    }

    public void create(String... args) {

        TimeMonitor.start();

        File dir = new File(config.getDataSrcFileDir());
        if (dir.isDirectory()) {
            File[] listFiles = dir.listFiles(new FilenameFilter() {
                @Override
                public boolean accept(File dir, String name) {
                    return name.endsWith("xlsx") || name.endsWith("xls");
                }
            });
            List<File> fileList = new ArrayList(listFiles.length);
            // 文件列表排序，因为文件中数据有依赖
            for (String sortedKey : fileNameSortedKey) {
                for (File file : listFiles) {
                    if (file.getName().endsWith(sortedKey + ".xlsx") || file.getName().endsWith(sortedKey + ".xls")) {
                        fileList.add(file);
                    }
                }
            }

            for (File file : fileList) {

                currentFileName = file.getName();

                // 读取excel文件数据
                List<Map<String, String>> srcData = readExcel(file.getAbsolutePath());

                if (srcData.size() == 0) {
                    continue;
                }

                // 每条线程执行的数据条数
                int perThreadExecuteNum = srcData.size()/config.getThreadNum();

                // 如果新增一个线程，那最后一个线程应该执行的数据条数
                int lastThreadExecuteNum = srcData.size() % config.getThreadNum();

                // 拆成多线程 如果当前数据条数能整除线程数，就按照当前线程个数，否则 +1
                config.setThreadNum((srcData.size() % config.getThreadNum() != 0) ? config.getThreadNum() + 1 : config.getThreadNum());

                CountDownLatch downLatch = new CountDownLatch(config.getThreadNum());

                log.info("文件：{} 执行线程数：{}", file.getName(), config.getThreadNum());

                for (int i = 0; i < config.getThreadNum(); i++) {
                    synchronized (this) {
                        int index = i;
                        new Thread(new Runnable() {
                            List<Map<String, String>> perThreadData = srcData.subList(index * perThreadExecuteNum, (((index + 1) * perThreadExecuteNum) > srcData.size()) || (index == config.getThreadNum() - 1) ? srcData.size() : (index + 1) * perThreadExecuteNum);
                            @Override
                            public void run() {
                                // 将数据填充模板文件后转成pdf并上传到dfs，得到fileId返回
                                createPdf(perThreadData, false);
                                // 生成sql文件
                                generateSqlFile(perThreadData, file);
                                downLatch.countDown();
                                log.info("第：{}个线程：{}执行：{}条数据已完成", index, Thread.currentThread().getName(), perThreadData.size());
                            }
                        }).start();
                    }

                }
                try {
                    downLatch.await();
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }

            }
        }

        TimeMonitor.end();
        System.exit(0);

    }

    private void generateSqlFile(List<Map<String, String>> srcData, File file) {
        // 每个文件中脚本的条数
        int countPerFile = 0;

        // 文件数
        int countFile = (srcData.size() % maxCountPerFile) == 0 ? (srcData.size() / maxCountPerFile) : (srcData.size() / maxCountPerFile) + 1;

        for (int i = 1; i <= countFile; i ++) {
            // 将脚本模板和数据组装成最终的脚本。指定数据源文件地址
            String sql = buildSql(sqlTemplateMapping.get(file.getName().substring(0, file.getName().lastIndexOf("."))), srcData.subList(maxCountPerFile * (i - 1) , (maxCountPerFile * i) <= srcData.size() ? (maxCountPerFile * i)  : srcData.size()));

            // 计算文件脚本条数
            countPerFile = ((maxCountPerFile * i) <= srcData.size() ? (maxCountPerFile * i) : srcData.size()) - maxCountPerFile * (i - 1);

            // 数据变更内容模板化
            String sqlContentTemplatePrefix = "set autocommit=0;\r\n" + "begin;\r\n" + "\r\n" + "-- 172.16.16.180/cmis\r\n"
                    + "-- 一共更新" + countPerFile + "条\r\n" + "-- " + file.getName().substring(0, file.getName().lastIndexOf(".")) + "\r\n";
            String sqlContentTemplateSuffix = "\r\ncommit;";
            if (isAppendSqlFix) {
                sql = sqlContentTemplatePrefix + sql + sqlContentTemplateSuffix;
            }
            String dataSaveFileName = config.getDataModSaveFilePathPrefix().replaceAll("SUMMARY", file.getName().substring(0, file.getName().lastIndexOf(".")))
                    .replace("_THREADNO", Thread.currentThread().getName()).replace("YYYYMMDD", new SimpleDateFormat("yyyyMMdd").format(new Date()));

            // 生成最终的数据变更脚本文件。指定文件存放路径
            writeToFile(sql, (srcData.size() > maxCountPerFile) ? (dataSaveFileName + "_" + i + getDataModSaveFilePathSuffix) : (dataSaveFileName + getDataModSaveFilePathSuffix));

            log.info("生成脚本文件：{}，一共：{}条脚本", ((srcData.size() > maxCountPerFile) ? (dataSaveFileName + "_" + i + getDataModSaveFilePathSuffix) : (dataSaveFileName + getDataModSaveFilePathSuffix)), countPerFile);
        }
    }

    private void createPdf(List<Map<String, String>> srcData, boolean requireName) {
        srcData.forEach(data -> {
            Map<String, String> map = new HashMap<>();
            String template = "";
            if (currentFileName.contains("放款凭证")) {
                template = config.getLoanContractTemp();
                map.put("tradeNo", data.get("交易号"));
                map.put("appleTime", data.get("申请时间"));
                map.put("name", data.get("收款方全称"));
                map.put("bankName", data.get("收款方开户单位"));
                map.put("account", data.get("收款方账号"));
                map.put("upperNum", data.get("大写金额"));
                map.put("lowerNum", data.get("金额小写"));
                map.put("successTime", data.get("成功时间"));
                map.put("receiptNo", data.get("交易号"));
                createPdf(map, requireName, template);
            } else if (currentFileName.contains("扣款凭证")) {
                template = config.getRepayContractTemp();
                map.put("tradeNo", data.get("交易号"));
                map.put("bankInfo", data.get("扣款账户信息"));
                map.put("amount", data.get("扣款金额"));
                map.put("successTime", data.get("成功时间"));
                map.put("createTime", data.get("创建时间"));
                lianlianCreatePdfVersion2(map);
            }
            String file = config.getCzGuaranteeContractPath2() + data.get("交易号") + ".pdf";
            log.info("最终上传的文件：{}", file);
            data.put("文件ID", upload(file));
        });
    }

    private void lianlianCreatePdfVersion2(Map<String, String> map) {
        if (map == null || map.size() == 0) {
            return;
        }
        for (Map.Entry<String, String> entry : map.entrySet()) {
            if (StringUtils.isEmpty(entry.getValue())) {
                System.out.println("参数存在空值：" + entry.getKey() + "，自动跳转" + JSONObject.toJSONString(map));
                return;
            }
        }
        // 合同模板地址
        String template = config.getRepayContractTemp();
        if (StringUtils.isEmpty(template)) {
            System.out.println("合同模板不存在/app/reconciliation/data/cz_guarantee_contract_temp.docx");
            return;
        }
        // 填充的合同文件
        String templateDocx = config.getCzGuaranteeContractPath() + map.get("tradeNo") + ".docx";
        String templatePdf = config.getCzGuaranteeContractPath2() + map.get("tradeNo") + ".pdf";

        try {
            com.aspose.words.Document document = new com.aspose.words.Document(template);
            map.forEach((key, val) -> {
                try {
                    document.getRange().replace(Pattern.compile(key),val);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            });
            document.save(templateDocx);
            AsposeUtil.docxToPdf(templateDocx, templatePdf);

        } catch (Exception e) {
            System.out.println("稠州担保合同生成异常:" + templateDocx);
            e.printStackTrace();
        } finally {

        }

    }

    private String upload(String filePath) {
        long start = System.currentTimeMillis();
        DFSMessage resp = DFSClient.instance(config.getMsjrDfsGatewayConfig()).save(filePath);
        long end = System.currentTimeMillis();
        if(resp.isSuccess()) {
            long elipsed = (end-start)/1000;
            log.info("文件：{}上传成功，返回文件ID：{}", filePath, resp.getFileId());
            return resp.getFileId();
        }else {
            log.error("文件：{}上传失败，原因：{}", filePath, resp.getMessage());
            return null;
        }
    }

    private void createPdf(Map<String, String> srcData, boolean requireName, String template) {
        if (srcData == null || srcData.size() == 0) {
            return;
        }
        for (Map.Entry<String, String> entry : srcData.entrySet()) {
            if (StringUtils.isEmpty(entry.getValue())) {
                log.warn("参数存在空值：{}，自动跳过：{}", entry.getKey(), JSONObject.toJSONString(srcData));
                return;
            }
        }
        // 合同模板地址

        String templateDocx;
        String templatePdf;
        if (requireName) {
            templateDocx = config.getCzGuaranteeContractPath() + srcData.get("name") + "_" + srcData.get("tradeNo") + ".docx";
            templatePdf = config.getCzGuaranteeContractPath2() + srcData.get("name") + "_" + srcData.get("tradeNo") + ".pdf";
        } else {
            templateDocx = config.getCzGuaranteeContractPath() + srcData.get("tradeNo") + ".docx";
            templatePdf = config.getCzGuaranteeContractPath2() + srcData.get("tradeNo") + ".pdf";
        }

        OPCPackage pack = null;
        XWPFDocument doc = null;
        FileOutputStream fopts = null;
        try {
            //获取docx解析对象
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(template));
            //解析替换表格对象
            changeTable(document, srcData);
            //生成新的word
            File file = new File(templateDocx);
            FileOutputStream stream = new FileOutputStream(file);
            document.write(stream);
            stream.close();
            log.info("生成word文档：{}", templateDocx);
            log.info("生成pdf文档：{}", templatePdf);
            AsposeUtil.docxToPdf(templateDocx, templatePdf);
        } catch (Exception e) {
            log.error("稠州担保合同生成异常: " + templateDocx, e);
        } finally {
            if (pack != null) {
                try {
                    pack.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (fopts != null) {
                try {
                    fopts.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void changeTable(XWPFDocument document, Map<String, String> textMap){
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        //只会有一个表格对象，所以取第一个表格对象
        List<XWPFTableRow> rows = tables.get(0).getRows();
        //遍历表格,并替换模板
        eachTable(rows, textMap);
    }



    /**
     * 遍历表格
     * @param rows 表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(List<XWPFTableRow> rows ,Map<String, String> textMap){
        for (XWPFTableRow row : rows) {
            //得到表格每一行的所有表格
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if(checkText(cell.getText())){
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            run.setText(changeValue(run.toString(), textMap),0);
                        }
                    }
                }
            }
        }
    }



    /**
     * 匹配传入信息集合与模板
     * @param value 模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static String changeValue(String value, Map<String, String> textMap){
        Set<Map.Entry<String, String>> textSets = textMap.entrySet();
        for (Map.Entry<String, String> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String key = "${"+textSet.getKey()+"}";
            if(value.indexOf(key)!= -1){
                value = textSet.getValue();
            }
        }
        //模板未匹配到区域替换为空
        if(checkText(value)){
            value = "";
        }
        return value;
    }


    /**
     * 判断文本中时候包含$
     * @param text 文本
     * @return 包含返回true,不包含返回false
     */
    public static boolean checkText(String text){
        boolean check  =  false;
        if(text.indexOf("$")!= -1){
            check = true;
        }
        return check;

    }

    /**
     * 写出脚本到指定的文件中
     * @param sql
     * @param filePath
     */
    public static void writeToFile(String sql, String filePath) {
        OutputStream out = null;
        try {
            File file = new File(filePath);
            if (!file.exists()) {
                if (!file.getParentFile().exists()) {
                    file.getParentFile().mkdirs();
                }
                file.createNewFile();
            } else {
                // 存在则先删除再新建
                file.delete();
                file.createNewFile();
            }
            out = new FileOutputStream(file);
            out.write(sql.getBytes());
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 根据脚本模板用数据填充生成数据变更内容
     * @param sqlTemplate
     * @param list
     * @return
     */
    public static String buildSql(String sqlTemplate, List<Map<String, String>> list) {
        StringBuilder sqlBuilder = new StringBuilder("");
        count = list.size();
        for (Map<String, String> map : list) {
            String tempStr = sqlTemplate;
            for (Map.Entry<String, String> entry : map.entrySet()) {
                if (Objects.isNull(entry.getKey())) {
                    log.info("停住");
                }

                if (Objects.isNull(entry.getValue())) {
                    log.info("{} 的值为空", entry.getKey());
                }
                tempStr = tempStr.replaceAll(entry.getKey(), entry.getValue());
            }
            sqlBuilder.append(tempStr).append("\r\n");
        }
        return sqlBuilder.toString();
    }

    /**
     * 读取excel
     * @param filePath
     * @return
     */
    public static List<Map<String, String>> readExcel(String filePath) {
        if (filePath.endsWith("xls")) {
            return readXLSExcel(filePath);
        } else if (filePath.endsWith("xlsx")) {
            return readXLSXExcel(filePath);
        }
        return null;
    }

    /**
     * 读取指定的excel文件
     * @param filePath
     * @return
     */
    public static List<Map<String, String>> readXLSExcel(String filePath) {
        // excel数据返回
        List<Map<String, String>> result = new ArrayList<Map<String, String>>();
        // 每行数据
        Map<String, String> rowData = null;
        try {
            // 1.获取文件输入流
            InputStream inputStream = new FileInputStream(filePath);

            // 如果是加密了得文件，做密码校验处理
//            POIFSFileSystem pfs = new POIFSFileSystem(inputStream);
//            inputStream.close();
//            EncryptionInfo encInfo = new EncryptionInfo(pfs);
//            Decryptor decryptor = Decryptor.getInstance(encInfo);
//            decryptor.verifyPassword("!Q@W3e4r");
//            HSSFWorkbook workbook = new HSSFWorkbook(decryptor.getDataStream(pfs));

            // 2.获取Excel工作簿对象
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            // 3.得到Excel工作表对象
            XSSFSheet sheetAt = workbook.getSheetAt(0);
            System.out.println("一共：" + sheetAt.getLastRowNum() + "行");
            // 4.循环读取表格数据

            // 用来处理数字类型转字符串
            DecimalFormat format = new DecimalFormat("###########.#####");


            DecimalFormat dateformat = new DecimalFormat("yyyy-mm-dd");

            // 列下标和列名容器
            Map<Integer, String> index2ColMapping = new HashMap<>();

            for (Row row : sheetAt) {
//                System.out.println("当前行数：" + row.getRowNum());
                // 首行（即表头）不读取
//                if (row.getRowNum() == 0) {
//                    continue;
//                }
                // 读取当前行中单元格数据，索引从0开始
                int cellNum = row.getLastCellNum();
                rowData = new HashMap<String, String>();
                for (int index = 0; index < cellNum; index ++) {
                    if (row.getRowNum() == 0) {
                        // 将列下标和表头名存在映射中
                        index2ColMapping.put(index, row.getCell(index).getStringCellValue());
                    } else {
                        if (Objects.isNull(row.getCell(index))) {
                            if (!Objects.isNull(index2ColMapping.get(index))) {
                                rowData.put(index2ColMapping.get(index), "");
                            }
                            continue;
                        }
                        String name = index2ColMapping.get(index);
                        if (CellType.STRING.compareTo(row.getCell(index).getCellType()) == 0) {
                            rowData.put(index2ColMapping.get(index), row.getCell(index).getStringCellValue());
                        } else if (CellType.NUMERIC.compareTo(row.getCell(index).getCellType()) == 0) {
                            if (DateUtil.isCellDateFormatted(row.getCell(index))) {
                                rowData.put(index2ColMapping.get(index), (String) rowFormatMapping.get(obtainFileName(filePath, ".xlsx")).get(index2ColMapping.get(index)).format(row.getCell(index).getDateCellValue()));
                            } else {
                                rowData.put(index2ColMapping.get(index), format.format(row.getCell(index).getNumericCellValue()));
                            }
                        } else if (CellType.FORMULA.compareTo(row.getCell(index).getCellType()) == 0) {
                            rowData.put(index2ColMapping.get(index), String.valueOf(row.getCell(index).getNumericCellValue()));
                        } else if (CellType.BLANK.compareTo(row.getCell(index).getCellType()) == 0) {
                            rowData.put(index2ColMapping.get(index), null);
                        } else {
                            log.info("类型是：{}", row.getCell(index).getCellType());
                        }
                    }
                    // 针对资产信息需求的定制化处理
                }
                if (rowData.size() > 0) {
                    // 针对用户信息特殊化处理
                    if (filePath.endsWith("用户信息.xlsx")) {
                        idNo2CustNo.put(rowData.get("借款人证件号"), rowData.get("用户编号"));
                        idNo2Name.put(rowData.get("借款人证件号"), rowData.get("借款人名称"));
                        rowData.put("证件类型", "身份证".equals(rowData.get("证件类型")) ? "1" : "2");
                        rowData.put("客户状态", "已认证".equals(rowData.get("客户状态")) ? "1" : "0");
                        if (Objects.isNull(rowData.get("联系地址"))) {
                            rowData.put("联系地址", "");
                        }
                        if (Objects.isNull(rowData.get("邮箱"))) {
                            rowData.put("邮箱", "");
                        }
                    } else if (filePath.endsWith("放款信息.xlsx")) {
                        rowData.put("用户编号", idNo2CustNo.get(rowData.get("借款人证件号")));
                        rowData.put("放款状态", "200");
                        rowData.put("贷款利率", new BigDecimal(rowData.get("贷款利率")).multiply(new BigDecimal(100)).setScale(2, RoundingMode.HALF_UP).toPlainString());
                        if (Objects.isNull(rowData.get("放款城市"))) {
                            rowData.put("放款城市", "");
                        }
                    } else if (filePath.endsWith("还款计划.xlsx")) {
                        if (Objects.isNull(rowData.get("息费减免"))) {
                            rowData.put("息费减免", "0");
                        }
                        if (Objects.isNull(rowData.get("实际还款日期"))) {
                            rowData.put("实际还款日期", "");
                        }
                        if (Objects.isNull(rowData.get("还款方式"))) {
                            rowData.put("还款方式", "");
                        }
                        if (Objects.isNull(rowData.get("逾期标志"))) {
                            rowData.put("逾期标志", "N");
                        }
                        if (Objects.isNull(rowData.get("结清标志"))) {
                            rowData.put("结清标志", "N");
                        }
                    }
                    result.add(rowData);
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 读取指定的excel文件
     * @param filePath
     * @return
     */
    public static List<Map<String, String>> readXLSXExcel(String filePath) {

        // excel数据返回
        List<Map<String, String>> result = new ArrayList<Map<String, String>>();
        // 每行数据
        Map<String, String> rowData = null;
        try {
            // 1.获取文件输入流
            InputStream inputStream = new FileInputStream(filePath);

            // 如果是加密了得文件，做密码校验处理
//            POIFSFileSystem pfs = new POIFSFileSystem(inputStream);
//            inputStream.close();
//            EncryptionInfo encInfo = new EncryptionInfo(pfs);
//            Decryptor decryptor = Decryptor.getInstance(encInfo);
//            decryptor.verifyPassword("!Q@W3e4r");
//            HSSFWorkbook workbook = new HSSFWorkbook(decryptor.getDataStream(pfs));

            // 2.获取Excel工作簿对象
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            // 3.得到Excel工作表对象
            XSSFSheet sheetAt = workbook.getSheetAt(0);
            System.out.println("一共：" + sheetAt.getLastRowNum() + "行");
            // 4.循环读取表格数据

            // 用来处理数字类型转字符串
            DecimalFormat format = new DecimalFormat("###########.#####");


            DecimalFormat dateformat = new DecimalFormat("yyyy-mm-dd");

            // 列下标和列名容器
            Map<Integer, String> index2ColMapping = new HashMap<>();

            for (Row row : sheetAt) {
//                System.out.println("当前行数：" + row.getRowNum());
                // 首行（即表头）不读取
//                if (row.getRowNum() == 0) {
//                    continue;
//                }
                // 读取当前行中单元格数据，索引从0开始
                int cellNum = row.getLastCellNum();
                rowData = new HashMap<String, String>();
                for (int index = 0; index < cellNum; index ++) {
                    if (row.getRowNum() == 0) {
                        // 将列下标和表头名存在映射中
                        index2ColMapping.put(index, row.getCell(index).getStringCellValue());
                    } else {
                        if (Objects.isNull(row.getCell(index))) {
                            if (!Objects.isNull(index2ColMapping.get(index))) {
                                rowData.put(index2ColMapping.get(index), "");
                            }
                            continue;
                        }
                        String name = index2ColMapping.get(index);
                        if (CellType.STRING.compareTo(row.getCell(index).getCellType()) == 0) {
                            rowData.put(index2ColMapping.get(index), row.getCell(index).getStringCellValue());
                        } else if (CellType.NUMERIC.compareTo(row.getCell(index).getCellType()) == 0) {
                            if (DateUtil.isCellDateFormatted(row.getCell(index))) {
                                rowData.put(index2ColMapping.get(index), (String) rowFormatMapping.get(obtainFileName(filePath, ".xlsx")).get(index2ColMapping.get(index)).format(row.getCell(index).getDateCellValue()));
                            } else {
                                rowData.put(index2ColMapping.get(index), format.format(row.getCell(index).getNumericCellValue()));
                            }
                        } else if (CellType.FORMULA.compareTo(row.getCell(index).getCellType()) == 0) {
                            rowData.put(index2ColMapping.get(index), String.valueOf(row.getCell(index).getNumericCellValue()));
                        } else if (CellType.BLANK.compareTo(row.getCell(index).getCellType()) == 0) {
                            rowData.put(index2ColMapping.get(index), null);
                        } else {
                            log.info("类型是：{}", row.getCell(index).getCellType());
                        }
                    }
                    // 针对资产信息需求的定制化处理
                }
                if (rowData.size() > 0) {
                    // 针对用户信息特殊化处理
                    if (filePath.endsWith("用户信息.xlsx")) {
                        idNo2CustNo.put(rowData.get("借款人证件号"), rowData.get("用户编号"));
                        idNo2Name.put(rowData.get("借款人证件号"), rowData.get("借款人名称"));
                        rowData.put("证件类型", "身份证".equals(rowData.get("证件类型")) ? "1" : "2");
                        rowData.put("客户状态", "已认证".equals(rowData.get("客户状态")) ? "1" : "0");
                        if (Objects.isNull(rowData.get("联系地址"))) {
                            rowData.put("联系地址", "");
                        }
                        if (Objects.isNull(rowData.get("邮箱"))) {
                            rowData.put("邮箱", "");
                        }
                    } else if (filePath.endsWith("放款信息.xlsx")) {
                        rowData.put("用户编号", idNo2CustNo.get(rowData.get("借款人证件号")));
                        rowData.put("放款状态", "200");
                        rowData.put("贷款利率", new BigDecimal(rowData.get("贷款利率")).multiply(new BigDecimal(100)).setScale(2, RoundingMode.HALF_UP).toPlainString());
                        if (Objects.isNull(rowData.get("放款城市"))) {
                            rowData.put("放款城市", "");
                        }
                    } else if (filePath.endsWith("还款计划.xlsx")) {
                        if (Objects.isNull(rowData.get("息费减免"))) {
                            rowData.put("息费减免", "0");
                        }
                        if (Objects.isNull(rowData.get("实际还款日期"))) {
                            rowData.put("实际还款日期", "");
                        }
                        if (Objects.isNull(rowData.get("还款方式"))) {
                            rowData.put("还款方式", "");
                        }
                        if (Objects.isNull(rowData.get("逾期标志"))) {
                            rowData.put("逾期标志", "N");
                        }
                        if (Objects.isNull(rowData.get("结清标志"))) {
                            rowData.put("结清标志", "N");
                        }
                    }
                    result.add(rowData);
                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * 根据文件路径字符串获取文件名
     *  如 private static String dataSrcFileDir = "E:\\文档\\项目文档\\资产信息\\demo.xlsx" 获取结果为demo
     *  注意：windows和linux 对于分隔符的区别
     * @param filePath 文件路径字符串
     * @param suffix 文件后缀
     * @return
     */
    private static String obtainFileName(String filePath, String suffix) {
        // 获取文件名
        String[] filePathSplitArr = filePath.split("\\\\");
        return filePathSplitArr[filePathSplitArr.length - 1].substring(0, filePathSplitArr[filePathSplitArr.length - 1].lastIndexOf(suffix));
    }
}
