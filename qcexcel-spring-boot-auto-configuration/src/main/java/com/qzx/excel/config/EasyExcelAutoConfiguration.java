package com.qzx.excel.config;

import com.qzx.excel.utils.SpringBootBeanUtil;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.autoconfigure.condition.ConditionalOnWebApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Configuration//申明这是一个配置类
@ConditionalOnWebApplication//引用启动器的项目是web应用此自动配置模块才生效
public class EasyExcelAutoConfiguration {

    /**
     * 一次批量导入数量控制，如果配置文件没有配置qc.excel.batchCount，则默认一次导入5条
     */
    @Value("${qc.excel.batchCount:5}")
    private Integer batchCount;

    @Bean
    public SpringBootBeanUtil getSpringBootBean(){
        return new SpringBootBeanUtil();
    }

    @Bean(name = "excelProperties")
    public ExcelProperties getExcelProperties(){
        ExcelProperties excelProperties = new ExcelProperties(){{
            setBatchCount(batchCount);
        }};
        return excelProperties;
    }

}
