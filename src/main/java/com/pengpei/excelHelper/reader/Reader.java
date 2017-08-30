package com.pengpei.excelHelper.reader;

import java.io.IOException;
import java.util.List;

/**
 * Created by pengpei on 2017/8/28.
 * 定义读取网络输入流，IO流中的Excel信息，并且转换成相应的实体
 */
public interface Reader<T> {
    List<T> read(Class<T> clazz) throws IOException, InstantiationException, IllegalAccessException;
}
