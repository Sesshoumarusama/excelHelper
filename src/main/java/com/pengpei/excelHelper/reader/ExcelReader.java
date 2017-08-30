package com.pengpei.excelHelper.reader;

import com.pengpei.excelHelper.annotation.Cell;
import com.pengpei.excelHelper.annotation.Excel;
import com.pengpei.excelHelper.util.CellUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Created by pengpei on 2017/8/29.
 */
public class ExcelReader<T> extends AbstractReader {
    private Workbook workbook;
    private ReadModel model;

    /**
     * 当readModel为LeftToRight时，key为行号（从0开始）
     * 当readModel为TopToBottom时，key为单元格的下标号（从0开始）
     */
    private Map<Integer, Field> fieldMap = new ConcurrentHashMap<Integer, Field>();

    private ExcelReader(BufferedInputStream bs, FileType fileType) throws IOException {
        super(bs, fileType);
        switch (fileType) {
            case XLS:
                workbook = new HSSFWorkbook(bs);
                break;
            case XLSX:
                workbook = new XSSFWorkbook(bs);
                break;
        }
    }

    private ExcelReader(BufferedInputStream bs, FileType fileType, ReadModel model) throws IOException {
        this(bs, fileType);
        this.model = model;
    }

    public List<T> read(Class clazz) throws IOException, InstantiationException, IllegalAccessException {
        List<T> list = null;
        Excel excel = (Excel) clazz.getAnnotation(Excel.class);
        if(model == null){
            model = excel.model();
        }
        initColumnFieldMap(clazz, model);
        try {
            switch (model) {
                case TopToBottom:
                    list = readFromTopToBottom(excel, clazz);
                    break;
                case LeftToRight:
                    list = readFromLeftToRight(excel, clazz);
                    break;
            }
            return list;
        } finally {
            if (workbook != null)
                workbook.close();
            if (bs != null)
                bs.close();
        }
    }

    private List<T> readFromTopToBottom(Excel excel, Class clazz) throws IllegalAccessException, InstantiationException {
        List<T> list = new ArrayList<T>();
        int beginRow = excel.beginRow();
        Sheet sheet;
        Row row;
        T obj;
        Field field;
        org.apache.poi.ss.usermodel.Cell cell;
        Object cellValue;
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            sheet = workbook.getSheetAt(sheetNum);
            for (int rowIndex = beginRow; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                row = sheet.getRow(rowIndex);
                if (row == null)
                    continue;
                obj = (T) clazz.newInstance();
                //单元格编号从0开始
                for (int i = 0; i < row.getLastCellNum(); i++) {
                    if (!fieldMap.containsKey(i))
                        continue;
                    cell = row.getCell(i);
                    if(cell == null || CellUtils.isBlank(cell))
                        continue;
                    field = fieldMap.get(i);
                    cellValue = CellUtils.parseCellValue(field, cell);
                    field.set(obj, cellValue);
                }
                list.add(obj);
            }
        }
        return list;
    }

    private List<T> readFromLeftToRight(Excel excel, Class clazz) throws IllegalAccessException, InstantiationException {
        List<T> list = new ArrayList<>();
        int beginColumn = excel.beginColumn();//数据开始所在的列
        T obj;
        org.apache.poi.ss.usermodel.Cell cell;
        Field field;
        Object cellValue;
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            Sheet sheet = workbook.getSheetAt(sheetNum);
            for (int rowIndex = 0; rowIndex < sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null)
                    continue;
                for (int i = beginColumn; i < row.getLastCellNum(); i++) {
                    cell = row.getCell(i);
                    if (cell == null || CellUtils.isBlank(cell))
                        continue;
                    obj = (i - beginColumn) < list.size() ? list.get(i - beginColumn) : null;
                    if (obj == null) {
                        obj = (T) clazz.newInstance();
                        list.add(obj);
                    }
                    field = fieldMap.get(rowIndex);
                    cellValue = CellUtils.parseCellValue(field, cell);
                    field.set(obj, cellValue);
                }
            }
        }
        return list;
    }


    private void initColumnFieldMap(Class clazz, ReadModel readModel) {
        Field[] declaredFields = clazz.getDeclaredFields();
        Field field;
        for (int i = 0; i < declaredFields.length; i++) {
            field = declaredFields[i];
            Cell cell = field.getAnnotation(com.pengpei.excelHelper.annotation.Cell.class);
            Integer n;
            switch (readModel) {
                case TopToBottom:
                    String columnNum = cell.columnNum();
                    if (!CellUtils.hasText(columnNum))
                        continue;
                    n = CellUtils.parseIntForColumnNum(columnNum);
                    break;
                case LeftToRight:
                    n = cell.rowNum() - 1;
                    break;
                default:
                    throw new IllegalArgumentException("无效的读取模式！");
            }
            if (fieldMap.containsKey(n))
                throw new IllegalArgumentException("不能为字段指定重复的列编号");
            field.setAccessible(true);
            fieldMap.put(n, field);
        }
    }

    public static class ReaderBuilder {
        private BufferedInputStream is;
        private FileType fileType;
        private ReadModel model;

        private ReaderBuilder() {
        }

        public static ReaderBuilder newBuilder() {
            ReaderBuilder builder = new ReaderBuilder();
            return builder;
        }

        public ReaderBuilder setSource(InputStream is, FileType fileType) {
            if (this.is != null)
                throw new IllegalArgumentException("已经设置了数据文件源！");
            if (this.fileType != null)
                throw new IllegalArgumentException("已经设置了文件类型!");
            this.is = new BufferedInputStream(is);
            this.fileType = fileType;
            return this;
        }

        public ReaderBuilder setSource(String filePath, FileType fileType) throws FileNotFoundException {
            InputStream is = new FileInputStream(filePath);
            return this.setSource(is, fileType);
        }

        /**
         * 该该做会使@Excel中设置的ReadModel无效
         * @param model
         * @return
         */
        public ReaderBuilder setReadModel(ReadModel model){
            this.model = model;
            return this;
        }

        public Reader build() throws IOException {
            if (is == null)
                throw new IllegalArgumentException("找不到读取的数据源！");
            if (fileType == null)
                throw new IllegalArgumentException("无法识别文件类型");
            Reader reader;
            if(model == null){
                reader = new ExcelReader(this.is, this.fileType);
            }else {
                reader = new ExcelReader(this.is, this.fileType, this.model);
            }
            return reader;
        }
    }
}
