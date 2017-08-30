package com.pengpei.excelHelper.reader;

import java.io.BufferedInputStream;

/**
 * Created by pengpei on 2017/8/29.
 */
public abstract class AbstractReader implements Reader{
    protected BufferedInputStream bs;
    private FileType fileType;

    protected AbstractReader(BufferedInputStream bs, FileType fileType){
        this.bs = bs;
        this.fileType = fileType;
    }
}
