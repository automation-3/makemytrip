package com.makemytrip.util;

import java.awt.List;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CommonFunctions {

	public CommonFunctions() {
		
	}
	
	public static File f;
	public static FileInputStream fp;
	public static XSSFWorkbook workbook;
	public static XSSFSheet sheet;
	public static Row row;
	public static Cell cell;

	public static synchronized int getRowNumber(String filePath, String sheetName, String testCaseID) {
		int rowNumber = -1;
		f = new File(filePath);
		int lastRowNumber;
		int lastCellNumber;
		String rowValue = "";
		try {
			fp = new FileInputStream(f);
			try {
				workbook = new XSSFWorkbook(fp);
				sheet = workbook.getSheet(sheetName);
				lastRowNumber = sheet.getLastRowNum();
				for (int i = 1; i <= lastRowNumber; i++) {
					lastCellNumber = sheet.getRow(i).getLastCellNum();
					for (int j = 0; j < lastCellNumber; j++) {
						switch (sheet.getRow(i).getCell(j).getCellTypeEnum()) {
						case NUMERIC:
							rowValue = NumberToTextConverter.toText(sheet.getRow(i).getCell(j).getNumericCellValue());
							break;

						case STRING:
							rowValue = sheet.getRow(i).getCell(j).getStringCellValue();
							break;

						case BLANK:
							rowValue = "";
							break;

						default:
							rowValue = "";
						}
						if (testCaseID.equals(rowValue))
							rowNumber = i + 1;
						break;
					}
					if (rowNumber != -1)
						break;
				}
			} catch (IOException e) {
				System.out.println(e);
			}

		} catch (FileNotFoundException e) {
			System.out.println(e);
		}
		return rowNumber;
	}

	public static void main(String[] args) {
		String filePath = System.getProperty("user.dir") + "\\testdata\\test.xlsx";

		System.out.println(getRowNumber(filePath, "Sheet1", "PP10"));
	}

}
