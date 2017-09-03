package com.pengpei.excelHelper.util;

import org.dom4j.Attribute;

import java.lang.reflect.Field;
import java.lang.reflect.Type;

/**
 * Created by pengpei on 2017/9/2.
 */
public class XMLParseUtils {

    public static Object getValue(Attribute attribute, Field field) {
        Object value;
        switch (field.getGenericType().getTypeName()) {
            case "java.lang.String":
                value = attribute.getValue();
                break;
            case "java.lang.Integer":
            case "java.lang.Short":
            case "int":
            case "short":
                value = Integer.parseInt(attribute.getValue());
                break;
            case "java.lang.Double":
            case "double":
                value = Double.parseDouble(attribute.getValue());
                break;
//            case "java.util.Date":
////                value = CellUtils.parseDate(cell);
//                break;
            case "java.lang.Boolean":
            case "boolean":
                value = Boolean.parseBoolean(attribute.getValue());
                break;
            default:
                throw new IllegalArgumentException("暂不支持的数据类型");
        }
        return value;
    }
}
