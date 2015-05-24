package com.importer;

import java.io.File;
import java.io.FileInputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.importer.annotation.ExcelCell;

public class ExcelImporter {
	public static List<Object> importFile(Class targetObjectClass, File file) {
		List<Object> objects = Lists.newArrayList();
		Workbook book = null;
		try {
			FileInputStream in = new FileInputStream(file);
			book = new XSSFWorkbook(in);
			Sheet sheet = book.getSheetAt(0);
			Iterator<Row> row = sheet.rowIterator();
			Map<Integer, String> headerMap = getHeaderMap(row);
			Map<String, Method> setterMap = getSetterMap(targetObjectClass);
			while (row.hasNext()) {
				Row rown = row.next();
				Iterator<Cell> cells = rown.cellIterator();
				Object object = targetObjectClass.newInstance();
				int cellIndex = 0;
				while (cells.hasNext()) {
					Cell cell = cells.next();
					String cellName = (String) headerMap.get(cellIndex);
					invokeSetterMethod(setterMap, object, cell, cellName);
					cellIndex = cellIndex + 1;
				}
				objects.add(object);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			IOUtils.closeQuietly(book);
		}
		return objects;
	}

	// TODO Need to set the value base on the field type(or add a need dataType
	// property in CellType annotation)
	private static void invokeSetterMethod(Map<String, Method> setterMap,
			Object object, Cell cell, String cellName)
			throws IllegalAccessException, InvocationTargetException {
		if (setterMap.containsKey(cellName)) {
			Method setMethod = (Method) setterMap.get(cellName);
			setMethod.invoke(object, cell.getStringCellValue());
		}
	}

	private static Map<Integer, String> getHeaderMap(Iterator<Row> row) {
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
				String setMethodName = "set"
						+ fieldName.substring(0, 1).toUpperCase()
						+ fieldName.substring(1);
				Method method = targetObjectClass.getMethod(setMethodName,
						field.getType());
				result.put(excel.columnName(), method);
			}
		}
		return result;
	}
}
