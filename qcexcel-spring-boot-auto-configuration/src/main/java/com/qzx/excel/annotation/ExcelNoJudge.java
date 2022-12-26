package com.qzx.excel.annotation;


import java.lang.annotation.*;

/**
 * 忽略校验标注此注解的属性
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelNoJudge {

    String value() default "";

}
