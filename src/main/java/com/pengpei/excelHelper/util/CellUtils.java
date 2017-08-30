package com.pengpei.excelHelper.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;
import java.util.Date;

/**
 * Created by pengpei on 2017/8/29.
 */
public class CellUtils {

    public static boolean hasText(String str){
        return str != null && !"".equals(str.trim());
    }

    /**
     * 将工作簿中的编号解析成整数编号
     * @param columnNum
     * @return
     */
    public static Integer parseIntForColumnNum(String columnNum) {
        if(columnNum == null)
            throw new NullPointerException("columnNum 不能为空！");
        if(isNumberStr(columnNum))
            return Integer.valueOf(columnNum);

        String str = columnNum.toUpperCase();
        int len = str.length();
        char c;
        int num = 0;
        for (int i = 0; i < len - 1; i++) {
            c = str.charAt(i);
            if(!isLetter(c))
                throw new IllegalArgumentException("无法解析列编号");
            num += Math.pow(26, (len - i - 1)) * (c - 'A' + 1);
        }
        num += str.charAt(len - 1) - 'A';
        return num;
    }

    /**
     * 判断一个字符串是不是一个数字字符串
     * 例如：“908”，
     * @param str
     * @return
     */
    public static boolean isNumberStr(String str){
        char c;
        for (int i = 0; i < str.length(); i++) {
            c = str.charAt(i);
            if(Character.isDigit(c)){
                continue;
            }
            return false;
        }
        return true;
    }

    /**
     * 判断资格字符是不是英文字母a-z,或者是A-Z
     * @return
     */
    public static boolean isLetter(char c){
        return (c > 'A' && c < 'Z') || (c > 'a' && c < 'z');
    }

    public static String joinStr(Object... str){
        if(str == null || str.length <= 2)
            throw new IllegalArgumentException("只需要两个待拼接的字符串");
        StringBuilder sb = new StringBuilder();
        for (Object o : str){
            sb.append(o);
        }
        return sb.toString();
    }

    public static int parseInt(Cell cell){
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType){
            case NUMERIC:
                String s = String.valueOf(cell.getNumericCellValue()).split("\\.")[0];
                return Integer.parseInt(s);
            case STRING:
                return Integer.parseInt(cell.getStringCellValue());
            default:
                throw new RuntimeException("无法解析Excel中的值！");
        }
    }

    public static double parseDouble(Cell cell){
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType){
            case NUMERIC:
                return cell.getNumericCellValue();
            case STRING:
                return Double.parseDouble(cell.getStringCellValue());
            default:
                throw new RuntimeException("无法解析Excel中的值！");
        }
    }

    public static Date parseDate(Cell cell){
        return cell.getDateCellValue();
    }

    public static boolean parseBoolean(Cell cell){
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType){
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case STRING:
                return Boolean.parseBoolean(cell.getStringCellValue());
            default:
                throw new RuntimeException("无法解析Excel中的值！");
        }
    }

    public static String parseString(Cell cell){
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType){
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getDateCellValue());
            default:
                throw new RuntimeException("无法解析Excel中的值！");
        }
    }

    public static Object parseCellValue(Field field, Cell cell) {
        Object cellValue;
        switch (field.getGenericType().getTypeName()) {
            case "java.lang.String":
                cellValue = CellUtils.parseString(cell);
                break;
            case "java.lang.Integer":
            case "java.lang.Short":
            case "int":
            case "short":
                cellValue = CellUtils.parseInt(cell);
                break;
            case "java.lang.Double":
            case "double":
                cellValue = CellUtils.parseDouble(cell);
                break;
            case "java.util.Date":
                cellValue = CellUtils.parseDate(cell);
                break;
            case "java.lang.Boolean":
            case "boolean":
                cellValue = CellUtils.parseBoolean(cell);
                break;
            default:
                throw new IllegalArgumentException("暂不支持的数据类型");
        }
        return cellValue;
    }

    public static boolean isBlank(Cell cell) {
        return CellType.BLANK.equals(cell.getCellTypeEnum());
    }
}
