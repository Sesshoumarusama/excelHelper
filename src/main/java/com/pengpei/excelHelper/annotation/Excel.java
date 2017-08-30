package com.pengpei.excelHelper.annotation;

import com.pengpei.excelHelper.reader.ReadModel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by pengpei on 2017/8/29.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {
    /**
     * 定义开始读取数据的行号
     * 行编号是从0开始的
     * 指定工作簿开始读取的行，在ReadModel.TopToBottom，
     * 默认表头为一行，需要去掉
     * @return
     */
    int beginRow() default 1;

    /**
     * 定义开始读取数据的列号
     * 默认表头为一列，需要去掉
     * 列编号是从0开始
     * 指定工作簿开始读取的列，在ReadModel.LeftToRight指定
     * @return
     */
    int beginColumn() default 1;

    /**
     * 指定工作簿的读取顺序
     * @return
     */
    ReadModel model() default ReadModel.TopToBottom;
}
