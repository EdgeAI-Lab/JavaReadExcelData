package com.fan.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

public class ExcelUtil {
	
	public static void main(String[] args){
		
		try {

			/*
			 * 读取.xls文件，导入hssf包
			 * 读取.xlsx文件，导入xssf包
			 * 读取以上两种格式的文件，导入ss包
			 * Excel(ss = hssf + xssf) - 来自java POI官网
			 * 
			 */
			
			Workbook wb = WorkbookFactory.create(new FileInputStream("MyExcel.xlsx"));
			
			Sheet sheet = wb.getSheetAt(0);
			
		    for (Row row : sheet) {
		        for (Cell cell : row) {
		        	System.out.println(cell.getRichStringCellValue().getString());
		        	
//		            CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
//		            System.out.print(cellRef.formatAsString());
//		            System.out.print(" - ");

//		            switch (cell.getCellType()) {
//		                case Cell.CELL_TYPE_STRING:
//		                    System.out.println(cell.getRichStringCellValue().getString());
//		                    break;
//		                case Cell.CELL_TYPE_NUMERIC:
//		                    if (DateUtil.isCellDateFormatted(cell)) {
//		                        System.out.println(cell.getDateCellValue());
//		                    } else {
//		                        System.out.println(cell.getNumericCellValue());
//		                    }
//		                    break;
//		                case Cell.CELL_TYPE_BOOLEAN:
//		                    System.out.println(cell.getBooleanCellValue());
//		                    break;
//		                case Cell.CELL_TYPE_FORMULA:
//		                    System.out.println(cell.getCellFormula());
//		                    break;
//		                default:
//		                    System.out.println();
//		            }
		        }
		    }
			
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		   
	}
}
