package com.example.spreadsheet;

public class MainClass {

	public static void main(String[] args) {

		ReadExcel rxl = new ReadExcel("C:\\Users\\adria\\eclipse- workspace Java\\spreadsheetdesign\\ExcelSheets\\TestData.xlsx"); //Creating a pointer/object to the ReadExcel class
		WriteToExcel wxl = new WriteToExcel("C:\\Users\\adria\\eclipse- workspace Java\\spreadsheetdesign\\ExcelSheets\\TestData.xlsx"); //Creating a pointer/object to the WriteExcel class
		
		System.out.println(rxl.getCellValue("B1"));
		wxl.setCellValue("A2", 25);
		System.out.println(rxl.getCellValue("A2"));
		wxl.setCellValue("A5", "A1+A2+A3+A4");
		System.out.println(rxl.getCellValue("A5"));
	}
}