package com.pengpei.excelHelper.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by pengpei on 2017/8/29.
 * 工作部中的单元格
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Cell {
    /**
     * 单元格所在列的编号, 从0开始
     * 可以写数字如“0”，“11”，也可以写excel上列的编号，如“A”,"AB"
     * @return
     */
    String columnNum() default "A";

    /**
     * 单元格所在的行的编号, 从0开始
     * @return
     */
    int rowNum() default 1;
}
