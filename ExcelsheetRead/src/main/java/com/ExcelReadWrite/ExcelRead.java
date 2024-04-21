package com.ExcelReadWrite;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelRead {
	public static void main(String[] args) throws Exception{
		DataFormatter df=new DataFormatter();
		FileInputStream fis=new FileInputStream("Book1.xlsx.");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet("Sheet1");
		
		int rows=sh.getLastRowNum();
		for(int i=0;i<=rows;i++) {// rows
			
			int clos=sh.getRow(i).getLastCellNum();
			
			for(int j=0;j<clos;j++) {  //cols
				
				Cell c=sh.getRow(i).getCell(j);
				System.out.print(df.formatCellValue(c)+ " " );
				
		}
			System.out.println();
	}
		getCelldata(2, 0);
				
  }
	public static String getCelldata(int row,int clo)throws Exception {
		DataFormatter df=new DataFormatter();
		FileInputStream fis=new FileInputStream("Book1.xlsx.");
		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet("Sheet1");
		
		
		return df.formatCellValue(sh.getRow(row).getCell(clo));
		
		
	}

}
