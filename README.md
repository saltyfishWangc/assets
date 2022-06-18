本功能实现如下：
1.将excel中的表格数据读取，根据给定的word模板填充数据，转换成pdf文件，最终将pdf上传，再生成对应的脚本文件
详细配置见application.properties，配置实体类见com.example.msyd.assets.config.Config

目录配置：
1.对应的excel数据文件上传到${dataSrcFileDir}指定的目录下，文件名是放款凭证.xlsx、扣款凭证.xlsx，目前代码中支持的是这两个。数据表格文件格式是第一列是中文列名，从第二行开始是数据。放款凭证的列名分别是：交易号、申请时间、收款方全程、收款方开户单位、收款方账号、大写金额、金额小写、成功时间。扣款凭证的列名分别是：交易号、扣款账户信息、扣款金额、成功时间、创建时间。列名顺序不要求按照这个，但是列名必须保持一致。如果有新增列名需要填充到模板文件中去或者生成脚本，需要修改代码
2.创建${czGuaranteeContractPath}、${czGuaranteeContractPath2}两个目录。${czGuaranteeContractPath}存放的是根据模板填充完数据后的word文档，${czGuaranteeContractPath2}存放的是填充完数据后的word文档转换的pdf文件。生产环境执行这两个目录如/home/excel2pdf/out/doc、/home/excel2pdf/out/pdf
3.上传temp04.docx、temp-lianlian2pdfVersion2.docx到/home/excel2pdf。其中temp04.docx是放款凭证的模板、temp-lianlian2pdfVersion2.docx是扣款凭证的模板
4.上传jar包到服务器，进入到jar所在目录执行 java -jar assets-0.0.1-SNAPSHOT.jar --spring.profiles.active=product --threadNum=5
参数说明：
spring.profiles.active=product 表示使用application-product.properties配置
threadNum=5 表示开启5个线程来处理。注意：代码里面的多线程处理是指处理同一个原数据文件，同一个原数据文件处理完后才会开始下一个原数据文件，同样也是开启对应数量的线程处理。
注：目前没有用到线程池，一个文件处理完后，线程销毁，下一个文件开始了又创建对应数量的线程，这块性能很差。后面要是有需要改用线程池
5.生成的脚本文件在${dataModSaveFilePathPrefix}目录下，生产环境如：/home/assets/ 
最终生成的文件名cmis_dml(172.16.16.180)_SUMMARY_THREADNO_YYYYMMDD格式说明如下：
SUMMARY：由当前数据的文件名替换
THREADNO：当前操作的线程名替换
YYYYMMDD：当前日期按yyyyMMdd格式转换后替换
eg:cmis_dml(172.16.16.180)_扣款凭证_Thread-73_20220613.sql