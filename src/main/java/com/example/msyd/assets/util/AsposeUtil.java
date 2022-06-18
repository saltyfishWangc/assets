package com.example.msyd.assets.util;

import com.aspose.words.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * @author : LCheng
 * @date : 2020-12-25 13:47
 * description : Aspose工具类
 */
public class AsposeUtil {

    /**
     * 加载license 用于破解 不生成水印
     *
     * @author LCheng
     * @date 2020/12/25 13:51
     */
    private static void getLicense() throws Exception {
        try (InputStream is = AsposeUtil.class.getClassLoader().getResourceAsStream("License.xml")) {
            License license = new License();
            license.setLicense(is);
        }
    }

    /**
     * word转pdf
     *
     * @param wordPath word文件保存的路径
     * @param pdfPath  转换后pdf文件保存的路径
     * @author LCheng
     * @date 2020/12/25 13:51
     */
    public static void wordToPdf(String wordPath, String pdfPath) throws Exception {
        getLicense();
        File file = new File(pdfPath);
        try (FileOutputStream os = new FileOutputStream(file)) {
            Document doc = new Document(wordPath);
            doc.save(os, SaveFormat.PDF);
        }
    }

    public static boolean docxToPdf(String inPath, String outPath) throws Exception {
        getLicense(); // 验证License 若不验证则转化出的pdf文档会有水印产生

        FileOutputStream os = null;
        try {
            long old = System.currentTimeMillis();
            File file = new File(outPath); // 新建一个空白pdf文档
            os = new FileOutputStream(file);
            Document document= new Document(inPath); // Address是将要被转化的word文档
            TableCollection tables = document.getFirstSection().getBody().getTables();
            for (Table table : tables) {
                RowCollection rows = table.getRows();
                table.setAllowAutoFit(false);
                for (Row row : rows) {
                    CellCollection cells = row.getCells();
                    for (Cell cell : cells) {
                        CellFormat cellFormat = cell.getCellFormat();
                        cellFormat.setFitText(false);
                        cellFormat.setWrapText(true);
                    }
                }
            }
            document.save(os, SaveFormat.PDF);// 全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF,
// EPUB, XPS, SWF 相互转换
            long now = System.currentTimeMillis();
            System.out.println("pdf转换成功，共耗时：" + ((now - old) / 1000.0) + "秒"); // 转化用时
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }finally {
            if (os != null) {
                try {
                    os.flush();
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return true;
    }

    public static void main(String[] args) throws Exception {
        wordToPdf("C:\\Users\\Dell\\Desktop\\服务协议书-稠州业务评审费_inst2.pdf.docx","C:\\Users\\Dell\\Desktop\\服务协议书-稠州.pdf");
    }
}
