package com.example.ExcelDataInsertion;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataInsertionApplication {

	public static void main(String[] args) throws IOException {
		FileOutputStream fos = new FileOutputStream("C:\\Users\\2485082\\OneDrive - Cognizant\\Desktop\\ExcelDataInsertion\\written.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Data");
		XSSFRow row = sheet.createRow(0);
		row.createCell(0).setCellValue("User Name");
		row.createCell(1).setCellValue("Password");
		row = sheet.createRow(1);
		row.createCell(0).setCellValue("selva@gmail.com");
		row.createCell(1).setCellValue(123456);
		row = sheet.createRow(2);
		row.createCell(0).setCellValue("ganesh@gmail.com");
		row.createCell(1).setCellValue(123456);
		workbook.write(fos);
		workbook.close();
		fos.close();
		System.out.println("File Created Successfully");
	}
}
