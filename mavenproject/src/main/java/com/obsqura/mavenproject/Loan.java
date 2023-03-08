package com.obsqura.mavenproject;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Loan {
	

	
	public static void main(String[] args) {
		ArrayList<Integer>list=new  ArrayList<Integer>();
		try {
			File file=new File("C:\\Users\\KPS\\Desktop\\obsqura\\loan.xlsx");
		FileInputStream fis=new FileInputStream(file);//converting excel to byte file
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet=wb.getSheetAt(0);
		Iterator<Row>i=sheet.iterator();
		while(i.hasNext()) 
		{
			Row row=i.next();
			Iterator<Cell>itr=row.cellIterator() ;
			while(itr.hasNext()) 
			{
				Cell cell=itr.next();
			switch(cell.getCellTypeEnum()) 
			{case STRING:
				System.out.println(cell.getStringCellValue()+"\t\t\t");
				break;
			case NUMERIC:
				System.out.println(cell.getNumericCellValue()+"\t\t\t");
				list.add((int)cell.getNumericCellValue());
				break;
				default:
				break;
			
			}}wb.close();
			System.out.println(" ");
					
				

		}}catch(Exception e) {
					e.printStackTrace();
				}System.out.println(list);
				System.out.println("total amount");
				int p=(list.get(0)*list.get(1)*list.get(2));
				System.out.println(p);
				
		}
		
		
		
		}