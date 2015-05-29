package com.importer;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.xmlbeans.impl.xb.xsdschema.FieldDocument.Field;
import org.apache.xmlbeans.impl.xb.xsdschema.ListDocument.List;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.importer.annotation.ExcelCell;

public class ExcelImporter {
    public static <T> List<T> importFile(Class<T> targetObjectClass, Sheet sheet) {
        List<T> objects = Lists.newArrayList();
        FileInputStream in = null;
        try {
            Iterator<Row> row = sheet.rowIterator();
            Map<Integer, String> headerMap = getHeaderMapFromExcel(row);
            Map<String, Method> setterMap = getSetterMap(targetObjectClass);
            Map<String, Class<?>> typeMap = getTypeMap(targetObjectClass);
            while (row.hasNext()) {
                Row rown = row.next();
                Iterator<Cell> cells = rown.cellIterator();
                T object = targetObjectClass.newInstance();
                int cellIndex = 0;
                while (cells.hasNext()) {
                    Cell cell = cells.next();
                    String cellName = headerMap.get(cellIndex);
                    invokeSetterMethod(setterMap, typeMap, object, cell, cellName);
                    cellIndex = cellIndex + 1;
                }
                objects.add(object);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(in);
        }
        return objects;
    }


    public static <T> void exportFile(List<T> objects, Class<T> targetObejctClass, Sheet sheet)
                    throws Exception {
        Map<Integer, String> headerMap = getHeaderMapFromObjectClass(targetObejctClass);
        Map<String, Method> getterMap = getGetterMap(targetObejctClass);
        Map<String, Class<?>> typeMap = getTypeMap(targetObejctClass);
        Row row = sheet.createRow((short) 0);
        for (int i = 0; i < headerMap.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            cell.setCellValue(headerMap.get(i));
        }

        int rowIndex = 1;
        for (T object : objects) {
            row = sheet.createRow(rowIndex++);
            for (int i = 0; i < headerMap.size(); i++) {
                Cell cell = row.createCell(i);
                String cellName = headerMap.get(i);
                Method method = getterMap.get(cellName);

                if (typeMap.get(cellName).toString().equals(Integer.class.toString())) {
                    cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                    cell.setCellValue((Integer) method.invoke(object));
                } else if (typeMap.get(cellName).toString().equals(Double.class.toString())) {
                    cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
                    cell.setCellValue((Double) method.invoke(object));
                } else {
                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                    cell.setCellValue((String) method.invoke(object));
                }
            }
        }

    }

    private static Map<String, Method> getGetterMap(Class targetObjectClass)
                    throws NoSuchMethodException, SecurityException {
        Map<String, Method> result = Maps.newHashMap();
        Field fields[] = targetObjectClass.getDeclaredFields();
        for (Field field : fields) {
            ExcelCell excel = field.getAnnotation(ExcelCell.class);
            if (excel != null) {
                String fieldName = field.getName();
                String getMethodName =
                                "get" + fieldName.substring(0, 1).toUpperCase()
                                                + fieldName.substring(1);
                Method method = targetObjectClass.getMethod(getMethodName);
                result.put(excel.name(), method);
            }
        }
        return result;
    }


    private static Map<Integer, String> getHeaderMapFromObjectClass(Class targetObjectClass) {
        Map<Integer, String> result = Maps.newHashMap();
        Field fields[] = targetObjectClass.getDeclaredFields();
        int i = 0;
        for (Field field : fields) {
            ExcelCell excel = field.getAnnotation(ExcelCell.class);
            if (excel != null) {
                result.put(i++, excel.name());
            }
        }
        return result;
    }


    // TODO Need to set the value base on the field type(or add a need dataType
    // property in CellType annotation)
    private static void invokeSetterMethod(Map<String, Method> setterMap,
                    Map<String, Class<?>> typeMap, Object object, Cell cell, String cellName)
                    throws IllegalAccessException, InvocationTargetException {
        if (setterMap.containsKey(cellName)) {
            Method setMethod = setterMap.get(cellName);
            if (typeMap.get(cellName).toString().equals(Integer.class.toString())) {
                setMethod.invoke(object, Double.valueOf(cell.getNumericCellValue()).intValue());
            } else if (typeMap.get(cellName).toString().equals(Double.class.toString())) {
                setMethod.invoke(object, cell.getNumericCellValue());
            } else {
                setMethod.invoke(object, cell.getStringCellValue());
            }

        }
    }

    private static Map<Integer, String> getHeaderMapFromExcel(Iterator<Row> row) {
        Row title = row.next();
        Iterator<Cell> cellTitle = title.cellIterator();
        Map<Integer, String> titlemap = new HashMap<Integer, String>();
        int titleIndex = 0;
        while (cellTitle.hasNext()) {
            Cell cell = cellTitle.next();
            String value = cell.getStringCellValue();
            titlemap.put(titleIndex, value);
            titleIndex = titleIndex + 1;
        }
        return titlemap;
    }

    @SuppressWarnings("unchecked")
    private static Map<String, Method> getSetterMap(Class targetObjectClass)
                    throws NoSuchMethodException, SecurityException {
        Map<String, Method> result = Maps.newHashMap();
        Field fields[] = targetObjectClass.getDeclaredFields();
        for (Field field : fields) {
            ExcelCell excel = field.getAnnotation(ExcelCell.class);
            if (excel != null) {
                String fieldName = field.getName();
                String setMethodName =
                                "set" + fieldName.substring(0, 1).toUpperCase()
                                                + fieldName.substring(1);
                Method method = targetObjectClass.getMethod(setMethodName, field.getType());
                result.put(excel.name(), method);
            }
        }
        return result;
    }


    @SuppressWarnings("unchecked")
    private static Map<String, Class<?>> getTypeMap(Class targetObjectClass)
                    throws NoSuchMethodException, SecurityException {
        Map<String, Class<?>> result = Maps.newHashMap();
        Field fields[] = targetObjectClass.getDeclaredFields();
        for (Field field : fields) {
            ExcelCell excel = field.getAnnotation(ExcelCell.class);
            if (excel != null) {
                result.put(excel.name(), field.getType());
            }
        }
        return result;
    }
}
