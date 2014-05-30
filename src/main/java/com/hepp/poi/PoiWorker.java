package com.hepp.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A simple class to demonstrate the use of the Apache POI library to read and
 * write Microsoft Excel files.
 * 
 * @author elmar.hepp@gmail.com
 *
 */
public class PoiWorker {

	public static void main(String[] args) {
		PoiWorker poiReader = new PoiWorker();
		poiReader.readXLS();
		poiReader.readXLSX();
		poiReader.writeXLSX();
	}

	/**
	 * Method to write some data into an Excel xlsx file
	 */
	private void writeXLSX() {
		System.out.println("Starting writeXLSX");
		String outputFileNameXLSX = "output1.xlsx";

		try {
			// create some data
			Map<Integer, Object[]> map = new HashMap<Integer, Object[]>();
			map.put(1,
					new Object[] { "No", "First Name", "Last Name", "Email", "Phone", "Married" });
			map.put(2, new Object[] { 1, "Elmar", "Hepp", "hepp@gmail.com", "0981 2321", true });
			map.put(3,
					new Object[] { 2, "Bettina", "Hupp", "bet.hupp@gmail.com", "0171 7777", true });
			map.put(4,
					new Object[] { 3, "Djinges", "Kaan", "djinges@kaan.com", "0888 1212", false });

			// create fileOutputStream for the excel file
			FileOutputStream file = new FileOutputStream(new File(outputFileNameXLSX));

			// create a blank workbook instance
			XSSFWorkbook workbook = new XSSFWorkbook();

			// create a blank first sheet
			XSSFSheet sheet = workbook.createSheet("New Sheet");

			// get data into the sheet
			Set<Integer> keys = map.keySet();
			int rowNumber = 0;
			for (Integer key : keys) {
				Row row = sheet.createRow(rowNumber++);
				Object[] array = map.get(key);
				int cellNumber = 0;
				for (Object object : array) {
					Cell cell = row.createCell(cellNumber++);
					if (object instanceof String)
						cell.setCellValue((String) object);
					else if (object instanceof Integer)
						cell.setCellValue((Integer) object);
					else if (object instanceof Boolean)
						cell.setCellValue((Boolean) object);
				}
			}

			// write data into the output file
			workbook.write(file);
			file.close();
		} catch (IOException e) {
			System.err.println("IOException at writer the file " + outputFileNameXLSX + ", "
					+ e.getLocalizedMessage());
		}
	}

	/**
	 * Method to read an Excel xslx file
	 */
	private void readXLSX() {
		System.out.println("Starting readXLSX");
		String excelFileNameXLSX = "sample1.xlsx";

		try {
			// create fileInputStream for the excel file
			FileInputStream file = new FileInputStream(new File(excelFileNameXLSX));

			// create a workbook instance
			XSSFWorkbook workbook = new XSSFWorkbook(file);

			// get the first sheet
			XSSFSheet sheet = workbook.getSheetAt(0);

			// get an iterator for the rows
			Iterator<Row> rowIterator = sheet.iterator();

			// iterate over all rows in sheet 0
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				Iterator<Cell> cellIterator = row.cellIterator();
				// iterate over all cells in the row
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.print("Boolean: " + cell.getBooleanCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print("Numeric: " + cell.getNumericCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print("String: " + cell.getStringCellValue() + "\t");
					}
				}
				System.out.println("");
			}

			file.close();
		} catch (FileNotFoundException e) {
			System.err.println("FileNotFoundException at reading " + excelFileNameXLSX + ", " + e);
		} catch (IOException e) {
			System.err.println("IOException at reading file " + excelFileNameXLSX + ", " + e);
		}
	}

	/**
	 * Method to read an Excel xls file
	 */
	private void readXLS() {
		System.out.println("Starting readXLS");
		String excelFileNameXLS = "sample1.xls";

		try {
			// create fileInputStream for the excel file
			FileInputStream file = new FileInputStream(new File(excelFileNameXLS));

			// create a workbook instance
			HSSFWorkbook workbook = new HSSFWorkbook(file);

			// get the first sheet
			HSSFSheet sheet = workbook.getSheetAt(0);

			// get an iterator for the rows
			Iterator<Row> rowIterator = sheet.iterator();

			// iterate over all rows in sheet 0
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();

				Iterator<Cell> cellIterator = row.cellIterator();
				// iterate over all cells in the row
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						System.out.print("Boolean: " + cell.getBooleanCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_NUMERIC:
						System.out.print("Numeric: " + cell.getNumericCellValue() + "\t");
						break;
					case Cell.CELL_TYPE_STRING:
						System.out.print("String: " + cell.getStringCellValue() + "\t");
					}
				}
				System.out.println("");
			}

			file.close();
		} catch (FileNotFoundException e) {
			System.err.println("FileNotFoundException at reading " + excelFileNameXLS + ", " + e);
		} catch (IOException e) {
			System.err.println("IOException at reading file " + excelFileNameXLS + ", " + e);
		}
	}

}
