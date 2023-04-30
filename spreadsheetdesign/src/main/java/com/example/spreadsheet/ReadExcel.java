package com.example.spreadsheet;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	
	String xlFilePath = "";
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	FileInputStream fis = null;
	
	public ReadExcel(String xlFilePath) {
		
		try
		{
			this.xlFilePath = xlFilePath;
			fis = new FileInputStream(new File(this.xlFilePath));
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		}
		catch(Exception e)
		{
			e.printStackTrace();
		}
	}
	public int getCellValue(String cellId)
	{
		CellAddress cellAddress = new CellAddress(cellId);
		Row row = sheet.getRow(cellAddress.getRow());
		Cell cell = row.getCell(cellAddress.getColumn());
		
		return (int)cell.getNumericCellValue();
		
		//In case, we need to add values other than Integers, we can use switch-case.
		
		/*switch(cell.getCellType()) {
			case BOOLEAN:
				System.out.println(cell.getBooleanCellValue());
				break;
			case NUMERIC:
				System.out.println(cell.getNumericCellValue());
				break;
			case STRING:
				System.out.println(cell.getStringCellValue());
				break;
			case FORMULA:
				System.out.println(cell.getNumericCellValue());
				break;
			default:
				System.out.println(cell.getStringCellValue());
				break;
		}*/
	}
}
