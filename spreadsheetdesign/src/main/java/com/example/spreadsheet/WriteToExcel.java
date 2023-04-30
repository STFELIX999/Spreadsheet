package com.example.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteToExcel {
	
	String xlFilePath = "";
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFCellStyle style;
	FileInputStream fis = null;
	FileOutputStream fos = null;
	FormulaEvaluator evaluator = null;
	
	public WriteToExcel(String xlFilePath) {
		
		try
		{
			this.xlFilePath = xlFilePath;
			fis = new FileInputStream(new File(this.xlFilePath));
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0); // Accessing the first sheet of the Excel Workbook
			evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	public void setCellValue(String cellId, String val){	//Used when we are passing arithmetic expressions as arguments
		
		CellAddress cellAddress = new CellAddress(cellId);
		Row row = sheet.getRow(cellAddress.getRow());
		Cell cell = row.getCell(cellAddress.getColumn());
		
		cell.setCellFormula(val);
		evaluator.evaluateFormulaCell(cell);
		if (cell.getCellType() == CellType.FORMULA) 
		{
			cell.setCellType(CellType.NUMERIC);
		}
		try
		{
			fos = new FileOutputStream(new File(this.xlFilePath));
			workbook.write(fos);
			fos.close();
			fis.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}	
	}
	
	public void setCellValue(String cellId, int val) {	// Used when we pass an Integer argument
		
		CellAddress cellAddress = new CellAddress(cellId);
		Row row = sheet.getRow(cellAddress.getRow());
		Cell cell = row.getCell(cellAddress.getColumn());
		cell.setCellValue(val);
		try
		{
			fos = new FileOutputStream(new File(this.xlFilePath));
			workbook.write(fos);
			fos.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
		
	}
}
