package com.obsqura.mavenproject;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Employee {

	public static void main(String[] args) {
		ArrayList<Integer>id=new ArrayList<Integer>();//declare arraylist
		try {
		File file=new File("C:\\Users\\KPS\\Desktop\\obsqura\\EMPLOYEE.xlsx");//call excel file
		FileInputStream fis=new FileInputStream(file);	// convert to byte file
         XSSFWorkbook wb=new XSSFWorkbook(fis);//call the converted file
         XSSFSheet sheet=wb.getSheetAt(1);//calling current sheet in the created workbook
         Iterator<Row>itr=sheet.iterator();//iterating row in the sheet to get the output
         while(itr.hasNext()) {
        	 Row row=itr.next();
         Iterator<Cell>i=row.cellIterator();//iterating the next cell to get the adjacent cells
         while(i.hasNext()) {
        	 Cell cell=i.next(); 
        	 switch(cell.getCellTypeEnum()){
        	 case STRING:
        		 System.out.println(cell.getStringCellValue()+"\t");
        		 break;
        	 case NUMERIC:
        		 System.out.println(cell.getNumericCellValue()+"\t");
        		 break;
        	 case BOOLEAN:
        		 System.out.println(cell.getBooleanCellValue()+"\t");
        		 break;
			default:
				break;
        		 
        	 }// to close switch
         }wb.close();// to close 2nd while loop//wb.close is optional
       System.out.println(" ");// it is provided to get space between the entry .it is optional
         
         
        	 
         }//to close first while loop
        	 
         
	}catch(Exception e) {//} to close try loop
e.printStackTrace();
}
		



		
	}}
