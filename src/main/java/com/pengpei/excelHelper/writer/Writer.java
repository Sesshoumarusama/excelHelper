package com.pengpei.excelHelper.writer;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * Created by pengpei on 2017/8/31.
 * 定义了写组件的基本行为，可以向文件或者是网络写数据
 */
public interface Writer<T> {
    void write(OutputStream os, List<T> dataList) throws Exception;

    void write(String filePath, List<T> dataList) throws Exception;
}
