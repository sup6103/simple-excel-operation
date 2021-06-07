package com.generic.task.rwexcel;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToPojo {

	public <T> List<T> convertExcelToPojo(InputStream iostream, Class<T> clazz) {

		List<T> listOfClass = new ArrayList<>();
		try {
			T newInstance = null;

			final DataFormatter formatter = new DataFormatter();
			final XSSFWorkbook workbook = new XSSFWorkbook(iostream);
			final XSSFSheet firstSheet = workbook.getSheetAt(0);
			final Iterator<Row> iterator = firstSheet.iterator();
			final Map<Integer, String> columnHeadingMap = new HashMap<>();
			String currentColumnHeading = "";

			while (iterator.hasNext()) {

				final Row currentRow = iterator.next();
				final int rowNum = currentRow.getRowNum();
				if (rowNum == 0) {

					final Iterator<Cell> cellIterator = currentRow.cellIterator();
					while (cellIterator.hasNext()) {
						final Cell cell = cellIterator.next();
						if (null != cell.getStringCellValue()
								&& !cell.getStringCellValue().replaceAll("\\s+", "").equals("")) {
							columnHeadingMap.put(cell.getColumnIndex(), cell.getStringCellValue());
						}
					}
				} else {
					final Iterator<Cell> cellIterator = currentRow.cellIterator();
					while (cellIterator.hasNext()) {
						newInstance = clazz.newInstance();
						final Cell cell = cellIterator.next();
						final int columnIndex = cell.getColumnIndex();
						currentColumnHeading = columnHeadingMap.get(columnIndex);
						formatter.formatCellValue(cell);
						Field f = null;
						if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
							f = newInstance.getClass().getDeclaredField(currentColumnHeading);
							f.setAccessible(true);
							f.set(newInstance, formatter.formatCellValue(cell));
							f.setAccessible(false);
						} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							f = newInstance.getClass().getDeclaredField(currentColumnHeading);
							f.setAccessible(true);
							f.setInt(newInstance, Integer.parseInt(formatter.formatCellValue(cell)));
							f.setAccessible(false);
						} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							f = newInstance.getClass().getDeclaredField(currentColumnHeading);
							f.setAccessible(true);
							f.set(newInstance, formatter.formatCellValue(cell));
						}
						listOfClass.add(newInstance);
					}
				}

			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return listOfClass;

	}
}
