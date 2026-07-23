package com.comcast.crm.generic.data.fileutility;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelUtility 
{
	public String getDataFromExcel(String sheetName, int rowNum, int celNum) throws Exception
	{
		FileInputStream fis = new FileInputStream("./testData/testSciptData.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		String data = wb.getSheet(sheetName).getRow(rowNum).getCell(celNum).getStringCellValue();
		wb.close();
		return data;
	}
	
	public int getRowCount(String sheetName) throws Exception 
	{
		FileInputStream fis = new FileInputStream("./testData/testSciptData.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet(sheetName);
		int rowCount = sh.getLastRowNum();
		wb.close();
		return rowCount;	
	}
	
	public void setDataIntoExcel(String sheetName, int rowNum, int celNum, String data) throws Exception
	{
		FileInputStream fis = new FileInputStream("./testData/testScriptData.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		wb.getSheet(sheetName).getRow(rowNum).createCell(celNum);
		
		FileOutputStream fos = new FileOutputStream("./testData/testScriptData.xlsx");
		wb.write(fos);
		wb.close();
		
	}
	
	public void setDataIntoExcel1(String sheetName, int rowNum, int cellNum, String data) throws Exception 
	{
	    FileInputStream fis = new FileInputStream("./testData/testScriptData.xlsx");
	    Workbook wb = WorkbookFactory.create(fis);

	    Sheet sheet = wb.getSheet(sheetName);

	    // Create row if not present
	    Row row = sheet.getRow(rowNum);
	    if(row == null) {
	        row = sheet.createRow(rowNum);
	    }

	    // Create cell if not present and set value
	    Cell cell = row.getCell(cellNum);
	    if(cell == null) {
	        cell = row.createCell(cellNum);
	    }
	    
	    cell.setCellValue(data);

	    fis.close();

	    FileOutputStream fos = new FileOutputStream("./testData/testScriptData.xlsx");
	    wb.write(fos);
	    wb.close();
	    fos.close();
	}
	
	
	public String[][] getMultipleRow_ColumnData(String sheetName) throws Exception 
	{
		FileInputStream fis = new FileInputStream("./testData/testSciptData.xlsx");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet(sheetName);

		int rowCount = sh.getLastRowNum();
		int colCount = sh.getRow(0).getLastCellNum();

		String[][] data = new String[rowCount][colCount];

		for (int i = 1; i <= rowCount; i++)   // skip header
		{
			Row row = sh.getRow(i);
			for (int j = 0; j < colCount; j++)
			{
				data[i - 1][j] = row.getCell(j).getStringCellValue();          
//				Excel row iteration starts from i = 1
//				But array indexing starts from 0
			}
		}

		wb.close();
		return data;
	}

	
}
