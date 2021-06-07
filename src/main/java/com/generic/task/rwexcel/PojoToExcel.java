package com.generic.task.rwexcel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.generic.task.rwexcel.annotation.ColumnName;

public class PojoToExcel {

	public XSSFWorkbook convertPojoToExcel(List<Class<?>> listOfClsObject) {
		
		List<Object> headerList= new ArrayList<>();
		List<List<Object>> sheetData= new ArrayList<>();
				
		List<Field> fields = Arrays.asList(listOfClsObject.get(0).getClass().getDeclaredFields()); 
		fields.stream().filter(f -> f.isAnnotationPresent(ColumnName.class)).forEach(field -> { 
			field.setAccessible(true);
			try { 
				headerList.add(field.getAnnotation(ColumnName.class).columnName()); 
			} catch (Exception e) {
				System.out.println("Exception in reading annotations"+e);
			}
			field.setAccessible(false);
		});
		sheetData.add(headerList);
		listOfClsObject.stream().forEach(record->{
			List<Field> field= Arrays.asList(record.getClass().getDeclaredFields()); 
			List<Object> rowData= new ArrayList<>();
			field.stream().forEach(f->{
				f.setAccessible(true);
				try {
				rowData.add(f.get(record));
				if(rowData.size() == field.size()) {
					sheetData.add(rowData);
				}
				}catch(Exception e) {
					System.out.println("Exception in reading field data"+e);
				}
				f.setAccessible(false);
			});
			
		});
		
		XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Result Sheet");
        int rowNum = 0;
        for(List<Object> rows : sheetData) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : rows) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }
        return workbook;
	}
}
