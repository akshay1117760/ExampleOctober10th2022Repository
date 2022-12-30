package com.sgtesting.ExcelPOI;


import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWRITEcontent {

	public static void main(String[] args) {
		readcontent();
	}

	private static void readcontent()
	{
		FileOutputStream fout=null;
		Workbook wb=null;
		Sheet sh=null;
		Row row=null;
		Cell cell=null;

		try
		{

			wb=new XSSFWorkbook();
			sh=wb.createSheet("Information");
			row=sh.createRow(0);
			cell=row.createCell(0);
			cell.setCellValue("USERNAME");

			cell=row.createCell(1);
			cell.setCellValue("PASSWORD");


			row=sh.createRow(1);
			cell=row.createCell(0);
			cell.setCellValue("admin");

			cell=row.createCell(1);
			cell.setCellValue("manager");

			fout=new FileOutputStream("C:\\EXCEL\\Assignment1.xlsx");
			wb.write(fout);
		}catch (Exception e) 
		{
			e.printStackTrace();
		}
		finally
		{
			try
			{
				fout.close();
				wb.close();
			}catch (Exception e) 
			{
				e.printStackTrace();
			}
		}

	}
}
