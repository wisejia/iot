package com.poseidon.excel;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Excel02 {
	public static void main(String[] args) {

		Workbook workbook = new HSSFWorkbook();//~2003 xls
		//Workbook workbook2 = new XSSFWorkbook();// 2003~ xlsx
		
		Sheet sheet1 = workbook.createSheet("시트만듬");
		
		Row row = null;
		Cell cell = null;
		
		row = sheet1.createRow(0);
		
		cell = row.createCell(0);
		cell.setCellValue("이름");
		cell = row.createCell(1);
		cell.setCellValue("나이");
		cell = row.createCell(2);
		cell.setCellValue("연도");
		
		Row row2 = sheet1.createRow(1);
		cell = row2.createCell(0);
		cell.setCellValue("홍길동");
		cell = row2.createCell(1);
		cell.setCellValue("30");//?
		cell = row2.createCell(2);
		cell.setCellValue("=now()");
		
		Row row3 = sheet1.createRow(2);
		cell = row3.createCell(0);
		cell.setCellValue("김길동");
		cell = row3.createCell(1);
		cell.setCellValue(30);//?
		cell = row3.createCell(2);
		cell.setCellValue("=now()");
		
		
		File file = new File("c:\\temp\\poiExcel.xls");
		try {
			FileOutputStream fileOut = new FileOutputStream(file);
			workbook.write(fileOut);
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		
		
	}
}