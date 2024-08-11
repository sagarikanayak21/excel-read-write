package com.sn.write;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws FileNotFoundException {
		excelWrite();
	}
	
	private static void excelWrite() throws FileNotFoundException {
		XSSFWorkbook work = new XSSFWorkbook();
		XSSFSheet sheet = work.createSheet();
		XSSFRow headerRow = sheet.createRow(0);
		headerRow.createCell(0).setCellValue("StudentId");
		headerRow.createCell(1).setCellValue("Name");
		headerRow.createCell(2).setCellValue("Address");
		
		XSSFRow dataRow = sheet.createRow(1);
		dataRow.createCell(0).setCellValue("USBM101");
		dataRow.createCell(1).setCellValue("Sagarika Nayak");
		dataRow.createCell(2).setCellValue("Bhadrak");
		
		FileOutputStream fileout;
		fileout = new FileOutputStream("student.xlsx");
		try {
			work.write(fileout);
			System.out.println("Excel File Created successfully.");
		}catch(IOException e) {
			e.printStackTrace();
		}
	}

}
