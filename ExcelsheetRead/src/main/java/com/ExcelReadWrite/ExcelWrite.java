package com.ExcelReadWrite;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelWrite {

	public static Cell c =null;

	public static void main(String[] args) throws Exception {

		FileInputStream fis=new FileInputStream("Book1.xlsx.");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet("Sheet1");

		//data write in the cell 7,4

		if(sh.getRow(7)==null) 
			c=sh.createRow(7).createCell(4);
		
		
		else {
			if(sh.getRow(7).getCell(4)==null) 
				sh.getRow(7).createCell(4);
			}
			c.setCellValue("TheKiranAcademy");
			FileOutputStream fos= new FileOutputStream("Book1.xlsx");
			wb.write(fos);
			wb.close();
			fos.close();
			
		}

	}


