package com.pengpei.excelHelper.template;

import com.alibaba.fastjson.JSON;
import com.pengpei.excelHelper.template.exception.ParseException;
import com.pengpei.excelHelper.util.XMLParseUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Attribute;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by pengpei on 2017/9/2.
 * 定义了Excel XML配置的模板引擎
 *
 */
public final class ExcelTemplateEngine {

    private static final String TEMPALET_PATH = "/template/template.xml";

    private Document document;

    private Sheet sheet = new Sheet();

    private Table table = new Table();

    private Row head = new Row();

    private List<Row> rows = null;

    public ExcelTemplateEngine() throws DocumentException {
        SAXReader saxReader = new SAXReader();
        document = saxReader.read(Object.class.getResourceAsStream(TEMPALET_PATH));
    }

    public void configure(Workbook workbook) throws Exception {
        //开始解析配置模板
        parseTemplate(document);
        int colIndex = head.getCells().size() - 1;
        int rowIndex = table.getHeight() - 1;
        //设置样式
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER_SELECTION);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setFillForegroundColor((short) 13);// 设置背景色
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        org.apache.poi.ss.usermodel.Sheet sheet = workbook.createSheet(this.sheet.getName());
        CellRangeAddress cellRangeAddress = new CellRangeAddress(0, rowIndex, 0, colIndex);
        sheet.addMergedRegion(cellRangeAddress);

        org.apache.poi.ss.usermodel.Row row = sheet.createRow(0);
        org.apache.poi.ss.usermodel.Cell cell = row.createCell(0);
        cell.setCellValue(table.getTitle());

        row = sheet.createRow(++rowIndex);
        int cellNum = head.getCells().size();
        for (int i = 0; i < cellNum; i++) {
            row.createCell(i).setCellValue(head.getCells().get(i).getText());
        }

        if(rows != null && rows.size() > 0){
            for (Row r : rows){
                row = sheet.createRow(++rowIndex);
                cellNum = r.getCells().size();
                for (int i = 0; i < cellNum; i++) {
                    row.createCell(i).setCellValue(r.getCells().get(i).getText());
                }
            }
        }

    }

    /**
     * 解析模板
     * @param document
     */
    private void parseTemplate(Document document) throws Exception {
        Element sheetElement = document.getRootElement();
        sheet = (Sheet) parseNode(sheetElement, Sheet.class);
        Element tableElement = sheetElement.element("table");
        table = (Table) parseNode(tableElement, Table.class);

        Cell cell;
        Element thead = tableElement.element("thead");
        Element tr = thead.element("tr");
        List<Element> ths = tr.elements("th");
        if(ths == null)
            throw new ParseException("无法解析");
        List<Cell> cells = new ArrayList<>();
        for (Element e : ths){
            cell = new Cell();
            cell.setText(e.getText());
            cells.add(cell);
        }
        head.setCells(cells);

        //解析初始化数据
        Element tbody = tableElement.element("tbody");
        List<Element> trs = tbody.elements("tr");
        if(trs != null){
            rows = new ArrayList<>(trs.size());
            Row r;
            for (Element tr1 : trs){
                r = new Row();
                List<Element> ths1 = tr1.elements("th");
                if(ths1 == null || ths1.size() == 0)
                    continue;
                //保证与投中设置的列相同
                cells = new ArrayList<>(head.getCells().size());
                for (int i = 0; i < head.getCells().size(); i++) {
                    cell = new Cell();
                    cell.setText(ths1.get(i).getText());
                    cells.add(cell);
                }
                r.setCells(cells);
                rows.add(r);
            }
        }

    }



    private Object parseNode(Element element, Class clazz) throws Exception{
        Field[] fields = clazz.getDeclaredFields();
        Field field;
        Object obj = clazz.getDeclaredConstructor(this.getClass()).newInstance(this);
        for (int i = 0; i < fields.length; i++) {
            Object value = null;
            field = fields[i];
            field.setAccessible(true);
            Attribute attribute = element.attribute(field.getName());
            if(attribute != null)
                value = XMLParseUtils.getValue(attribute, field);
            if(value != null && !"".equals(value.toString()))
                field.set(obj, value);
        }
        return obj;
    }

    private class Sheet{
        //默认为
        private String name = "sheet1";

        public Sheet() {
        }

        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }
    }

    private class Table {
        private String title;
        private int height = 1;

        public Table() {
        }

        public String getTitle() {
            return title;
        }

        public void setTitle(String title) {
            this.title = title;
        }

        public int getHeight() {
            return height;
        }

        public void setHeight(int height) {
            this.height = height;
        }
    }


    public class Row{
        public Row() {
        }

        private List<Cell> cells;

        //默认占一个单元格的高度
        private int height = 1;

        public List<Cell> getCells() {
            return cells;
        }

        public void setCells(List<Cell> cells) {
            this.cells = cells;
        }

        public int getHeight() {
            return height;
        }

        public void setHeight(int height) {
            this.height = height;
        }
    }

    public class Cell {
        private String text;

        public Cell() {
        }

        public String getText() {
            return text;
        }

        public void setText(String text) {
            this.text = text;
        }
    }






    public static void main(String[] args) throws Exception {
        ExcelTemplateEngine engine = new ExcelTemplateEngine();
        engine.configure(new XSSFWorkbook());
        System.out.println(JSON.toJSONString(engine.sheet));
        System.out.println(JSON.toJSONString(engine.head));
        System.out.println(JSON.toJSONString(engine.table));
        System.out.println(JSON.toJSONString(engine.rows));
    }
}
