package com.example.spreadsheet;

import java.io.File;
import java.io.FileOutputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateExcel {

	public static void main(String[] args) throws Exception{
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet();
		sheet.createRow(0);
		sheet.getRow(0).createCell(0).setCellValue("First Name");
		sheet.getRow(0).createCell(1).setCellValue("Last Name");
		
		sheet.createRow(1);
		sheet.getRow(1).createCell(0).setCellValue("Manju");
		sheet.getRow(1).createCell(1).setCellValue("Prasad");
		
		sheet.createRow(2);
		sheet.getRow(2).createCell(0).setCellValue("Reji");
		sheet.getRow(2).createCell(1).setCellValue("Mathew");
		
		File file = new File("C:\\Users\\adria\\eclipse- workspace Java\\spreadsheetdesign\\ExcelSheets\\Test4.xlsx");	// Give the File Path here.
		FileOutputStream fos = new FileOutputStream(file);
		workbook.write(fos);
		fos.close();
		workbook.close();
	}
}
