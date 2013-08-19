package com.gs.test;

import static org.junit.Assert.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;


public class testExcel {
	public String fileToBeWrite = "D:\\123.xls";

	@Test
	public void test1() { //退休
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileToBeWrite));
			HSSFSheet sheet = workbook.getSheet("职工");
			FileOutputStream fileOut;
			fileOut = new FileOutputStream(fileToBeWrite);
			for (int j = 6; j <= 2286; j++) {
				HSSFRow row = sheet.getRow(j);
				//HSSFCell cell;
				String s = row.getCell(5).getStringCellValue();
				String r = s.substring(0, 4)+"-"+s.substring(4, 6)+"-"+s.substring(6, 8);
				System.out.println(row.getCell(0).getNumericCellValue()+row.getCell(1).getStringCellValue()+"  "+row.getCell(5).getStringCellValue()+"-----"+r);
				row.getCell(5).setCellValue(r);
			}
			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} 

	}
	
	@Test
	public void test2(){ //
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileToBeWrite));
			HSSFSheet sheet = workbook.getSheet("离休");
			FileOutputStream fileOut;
			fileOut = new FileOutputStream(fileToBeWrite);
			for (int j = 8; j <= 81; j++) {
				HSSFRow row = sheet.getRow(j);
				String s = row.getCell(5).getStringCellValue();
				String r = s.substring(0, 4)+"-"+s.substring(4, 6)+"-"+s.substring(6, 8);
				System.out.println(row.getCell(0).getNumericCellValue()+row.getCell(1).getStringCellValue()+"  "+row.getCell(5).getStringCellValue()+"-----"+r);
				row.getCell(5).setCellValue(r);
			}
			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} 
	}
	
	@Test
	public void test3(){ //
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileToBeWrite));
			HSSFSheet sheet = workbook.getSheet("退休");
			//FileOutputStream fileOut;
			//fileOut = new FileOutputStream(fileToBeWrite);
			for (int j = 7; j <= 1628; j++) {
				HSSFRow row = sheet.getRow(j);
				double s = row.getCell(5).getNumericCellValue();
				System.out.println(s);
				System.out.println(j+"==="+row.getCell(0).getStringCellValue()+row.getCell(1).getStringCellValue()+"  "+row.getCell(5).getStringCellValue());
				//row.getCell(5).setCellValue(r);
			}
			//workbook.write(fileOut);
			//fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} 
	}
	
	@Test
	public void test4(){ //
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(fileToBeWrite));
			HSSFSheet sheet = workbook.getSheet("供养");
			FileOutputStream fileOut;
			fileOut = new FileOutputStream(fileToBeWrite);
			for (int j = 5; j <= 315; j++) {
				HSSFRow row = sheet.getRow(j);
				String s = row.getCell(2).getStringCellValue();
				String r = s.substring(0, 4)+"-"+s.substring(4, 6)+"-"+s.substring(6, 8);
				System.out.println(row.getCell(1).getStringCellValue()+"   "+row.getCell(2).getStringCellValue()+"   "+r);
				//System.out.println(j+"==="+row.getCell(0).getStringCellValue()+row.getCell(1).getStringCellValue()+"  "+row.getCell(5).getStringCellValue());
				row.getCell(2).setCellValue(r);
			}
			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} 
	}

}
