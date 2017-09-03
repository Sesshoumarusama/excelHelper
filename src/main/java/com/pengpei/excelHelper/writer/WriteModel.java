package com.pengpei.excelHelper.writer;

/**
 * Created by pengpei on 2017/8/31.
 */
public enum WriteModel {
    /**
     * 向文件中追加数据
     */
    Append,
    /**
     * 清空文件中的数据，再添加
     */
    Truncate,
    /**
     * 先删除同名的文件，再创建
     */
    DropAndCreate;
}
