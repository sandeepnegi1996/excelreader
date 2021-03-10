package com.workon.Main;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.workon.utilities.excelReader.ExcelReader;

public class QueryFormatter {

	public static void main(String[] args) {
	    try {
	        FileInputStream file = new FileInputStream(new File("C:\\END1COB\\Test_101\\excelReader\\src\\main\\resources\\Book1.xlsx"));

	        //Create Workbook instance holding reference to .xlsx file
	        XSSFWorkbook workbook = new XSSFWorkbook(file);

	        //Get first/desired sheet from the workbook
	        XSSFSheet sheet = workbook.getSheetAt(0);

	        //Iterate through each rows one by one
	        
	        //this counter will be used as Id in the table
	        
	        int counter=1; 
	        Iterator<Row> rowIterator = sheet.iterator();
	        while (rowIterator.hasNext())
	        {
	            Row row = rowIterator.next();
	            //For each row, iterate through all the columns
	            Iterator<Cell> cellIterator = row.cellIterator();

	            while (cellIterator.hasNext()) 
	            {
	                Cell cell = cellIterator.next();
	                //Check the cell type and format accordingly
	                switch (cell.getCellType()) 
	                {
	                    case Cell.CELL_TYPE_NUMERIC:
	                        System.out.print(cell.getNumericCellValue() + "\t");
	                        break;
	                    case Cell.CELL_TYPE_STRING:
	                        //System.out.print(cell.getStringCellValue() + "\t");
	                        System.out.println("Insert into SECURITYCLASS values ("+counter+",'"+cell.getStringCellValue()+"'"   
	                        		+",' ',' ',' ')"
	                        		);
	                        
	                        counter++;
	                        break;
	                }
	            }
	            System.out.println("");
	        }
	        file.close();
	    } catch (Exception e) {
	        e.printStackTrace();
	    }
	}
	
}