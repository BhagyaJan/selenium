package org.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {
	
public static void main(String[] args) throws IOException {
	File fRead=new File("D:\\ExcelRead.xlsx");
	FileInputStream fis = new FileInputStream(fRead);
	Workbook wb = new XSSFWorkbook(fis);
	Sheet sheet = wb.getSheet("Sheet1");
	//System.out.println(sheet.getRow(2).getCell(1));
	System.out.println(sheet.getPhysicalNumberOfRows());
	System.out.println(sheet.getRow(2).getPhysicalNumberOfCells());
	
//	for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
//		Row r = sheet.getRow(i);
//		for(int j=0;j<r.getPhysicalNumberOfCells();i++)
//		{
//			Cell c= r.getCell(j);
//			System.out.println(c);
//		}
//	}
	
	//Iterator<Row> itr = sheet.iterator();  
	Iterator<Row> itr = sheet.rowIterator();
	
	SimpleDateFormat df = new SimpleDateFormat("dd-MMM-yyyy");
	while (itr.hasNext())                 
	{  
	Row row = itr.next();  
	Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column  
	while (cellIterator.hasNext())   
	{  
	Cell cell = cellIterator.next();  
	switch (cell.getCellType())               
	{  
	case Cell.CELL_TYPE_STRING:    //field that represents string cell type  
	System.out.println(cell.getStringCellValue());  
	break;  
	case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type  
		   if (DateUtil.isCellDateFormatted(cell)) {
               System.out.println(df.format(cell.getDateCellValue()));
           } else {
               System.out.println(cell.getNumericCellValue());
           }  
	break;  
	default:  
	}  
	}  
	System.out.println("");  
	}  
}

}
