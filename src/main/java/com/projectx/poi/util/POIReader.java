package com.projectx.poi.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class POIReader {
	
	public POIReader(){
		
	}
	public List<Object> getDataPerSheet(String folder, String fileName, int sheetNumber){
		
		List<Object> dataList = null; 
		
		try {
			
			dataList = new ArrayList<Object>();
			
		    FileInputStream file = new FileInputStream(new File("C:\\"+fileName));
		     
		    //Get the workbook instance for XLS file 
		    HSSFWorkbook workbook = new HSSFWorkbook(file);
		 
		    //Get first sheet from the workbook
		    HSSFSheet sheet = workbook.getSheetAt(sheetNumber);
		     
		    //Iterate through each rows from first sheet
		    Iterator<Row> rowIterator = sheet.iterator();
		    while(rowIterator.hasNext()) {
		    	
		        Row row = rowIterator.next();
		        
		        List<String> rowList = new ArrayList<String>(); 
		         
		        //For each row, iterate through each columns
		        Iterator<Cell> cellIterator = row.cellIterator();
		        while(cellIterator.hasNext()) {
		             
		            Cell cell = cellIterator.next();
		             
		            switch(cell.getCellType()) {
		                case Cell.CELL_TYPE_BOOLEAN:
		                	rowList.add(new Boolean(cell.getBooleanCellValue()).toString());
		                    break;
		                case Cell.CELL_TYPE_NUMERIC:
		                	rowList.add(new Double(cell.getNumericCellValue()).toString());
		                    break;
		                case Cell.CELL_TYPE_STRING:
		                	rowList.add(cell.getStringCellValue());
		                    break;
		            }
		        }
		        dataList.add(rowList);
		    }
		    file.close();
		    
		} catch (FileNotFoundException fileNotFoundException) {
			fileNotFoundException.printStackTrace();
		} catch (IOException ioException) {
			ioException.printStackTrace();
		}
		return dataList;
	}
}


