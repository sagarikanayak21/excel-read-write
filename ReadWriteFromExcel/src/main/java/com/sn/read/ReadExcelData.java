package com.sn.read;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData {
	public static void main(String[] args) {
		readExcel("D:\\workspace\\xlSheet\\ReadFromExcel\\Employee.xlsx");
	}
	private static void readExcel(String file) {
		try {
			XSSFWorkbook work = new XSSFWorkbook(new FileInputStream(file));
			
			XSSFSheet sheet = work.getSheet("Sheet1");
			XSSFRow row = null;
			int i=1;
			while((row = sheet.getRow(i)) != null) {
				System.out.println("Emp Id: " + (int) row.getCell(0).getNumericCellValue());
				System.out.println("First Name: " +row.getCell(1));
				i++;
			}
		}catch(IOException e) {
			e.printStackTrace();
		}
	}
}
