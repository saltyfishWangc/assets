package com.example.msyd.assets.config;

import lombok.Data;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.PropertySource;

/**
 * @author
 * @Description:
 * @date 2022/6/17 18:02
 */
@Data
//@PropertySource(value="application.yml", encoding = "UTF-8", ignoreResourceNotFound = false)
@Configuration
public class Config {

    /*数据源文件地址*/
    @Value("${dataSrcFileDir}")
    private String dataSrcFileDir;

    /*// 最终生成的数据变更文件存放地址 前缀*/
    @Value("${dataModSaveFilePathPrefix}")
    private String dataModSaveFilePathPrefix;

    /*临时模板文件*/
    @Value("${loanContractTemp}")
    private String loanContractTemp;

    @Value("${repayContractTemp}")
    private String repayContractTemp;

    @Value("${czGuaranteeContractPath}")
    private String czGuaranteeContractPath;

    @Value("${czGuaranteeContractPath2}")
    private String czGuaranteeContractPath2;

    @Value("${msjr-dfs-gateway-config}")
    private String msjrDfsGatewayConfig;

    @Value("${threadNum}")
    private Integer threadNum;

}
