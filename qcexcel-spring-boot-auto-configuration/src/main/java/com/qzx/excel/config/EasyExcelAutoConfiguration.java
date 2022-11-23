package com.qzx.excel.config;

import com.qzx.excel.utils.SpringBootBeanUtil;
import org.springframework.boot.autoconfigure.condition.ConditionalOnWebApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;


@Configuration//申明这是一个配置类
@ConditionalOnWebApplication//引用启动器的项目是web应用此自动配置模块才生效
public class EasyExcelAutoConfiguration {

    @Bean
    public SpringBootBeanUtil getSpringBootBean(){
        return new SpringBootBeanUtil();
    }

}
