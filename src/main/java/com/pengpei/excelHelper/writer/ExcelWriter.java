package com.pengpei.excelHelper.writer;

import com.pengpei.excelHelper.annotation.Excel;
import com.pengpei.excelHelper.reader.ExcelReader;
import com.pengpei.excelHelper.reader.FileType;
import com.pengpei.excelHelper.reader.ReadModel;
import com.pengpei.excelHelper.template.ExcelTemplateEngine;
import com.pengpei.excelHelper.util.CellUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Created by pengpei on 2017/8/31.
 */
public final class ExcelWriter<T> implements Writer<T> {
    private FileType fileType;
    private ReadModel readModel;
    private Workbook workbook;
    private ExcelTemplateEngine excelTemplateEngine;

    private Map<Integer, Field> fieldMap = new ConcurrentHashMap<>();

    private ExcelWriter(FileType fileType) throws Exception {
        this.fileType = fileType;
        //读取配置的模板
        createWorkbook(fileType);
        excelTemplateEngine = new ExcelTemplateEngine();
        excelTemplateEngine.configure(workbook);
    }

    private ExcelWriter(FileType fileType, ReadModel readModel) throws Exception {
        this(fileType);
        this.readModel = readModel;
    }

    private void createWorkbook(FileType fileType) {
        switch (fileType) {
            case XLSX:
                workbook = new XSSFWorkbook();
                break;
            case XLS:
                workbook = new HSSFWorkbook();
                break;
            default:
                throw new IllegalArgumentException("不支持的文件类型");
        }
    }

    @Override
    public void write(OutputStream os, List<T> dataList) throws Exception {
        try {
            if (os == null)
                throw new IllegalArgumentException("找不到输出目标！");
            write(dataList);
            workbook.write(os);
        } finally {
            if (workbook != null)
                workbook.close();
        }
    }

    @Override
    public void write(String filePath, List<T> dataList) throws Exception {
        try {
            File file = new File(filePath);
            if (!file.exists()) {
                file.createNewFile();
            }
            OutputStream os = new BufferedOutputStream(new FileOutputStream(file));
            write(dataList);
            workbook.write(os);
        } finally {
            if (workbook != null)
                workbook.close();
        }
    }

    private void write(List<T> dataList) throws Exception{
        if(dataList == null)
            return;
        
        Sheet sheet = workbook.getSheetAt(0);
        int lastRowNum = sheet.getLastRowNum();

        Class clzss = dataList.get(0).getClass();
        if(this.readModel == null){
            Excel excel = (Excel) clzss.getAnnotation(Excel.class);
            readModel = excel.model();
        }

        fieldMap.putAll(ExcelReader.initColumnFieldMap(clzss, readModel));

        Field field;
        Row row;
        Cell cell;
        for (T t : dataList){
            row = sheet.createRow(++lastRowNum);
            for (Integer i : fieldMap.keySet()){
                field = fieldMap.get(i);
                cell = row.createCell(i);
                CellUtils.setCellValue(field, cell, t);
            }
        }
    }

    public static class WriterBuilder {
        private FileType fileType;
        private ReadModel readModel;

        public static WriterBuilder newWriter() {
            return new WriterBuilder();
        }

        public WriterBuilder setFileType(FileType fileType) {
            this.fileType = fileType;
            return this;
        }

        public WriterBuilder setModel(ReadModel readModel) {
            this.readModel = readModel;
            return this;
        }

        public Writer build() throws Exception {
            if (fileType == null)
                throw new IllegalArgumentException("无法识别文件类型");
            if (readModel == null)
                return new ExcelWriter(fileType);
            return new ExcelWriter(fileType, readModel);
        }
    }
}
