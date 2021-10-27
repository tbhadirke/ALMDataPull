package org.apache.maven.DBConnect;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils1 

{

	private static XSSFSheet ExcelWSheet;
	private static XSSFWorkbook ExcelWbook;
	private static XSSFCell Cell;
	
	public static void SetExcelFile(String Path, String SheetName) throws IOException
	{
		FileInputStream ExcelFile = new FileInputStream(Path);
		ExcelWbook = new XSSFWorkbook(ExcelFile);
		ExcelWSheet = ExcelWbook.getSheet(SheetName); 
	}
	
	public static String getCellData(int RowNum, int ColNum)
	{
	Cell = ExcelWSheet.getRow(RowNum).getCell(ColNum);
	String CellData = Cell.getStringCellValue();
	return CellData;
	}

	public static int getTotalRows() {
		int i= ExcelWSheet.getLastRowNum();
		return i;
	}
}
