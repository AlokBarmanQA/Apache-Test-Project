package com.demo.writeExcel;

import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateInvoice {

	public static void main(String[] args) {

		try {
			Workbook workbook = new XSSFWorkbook();
			Sheet sh = workbook.createSheet("Invoices");
			String[] columnHeadings = {"Item id", "Item Name", "Qty", "Item Price", "Sold Date"};
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeight((short)12);
			headerFont.setColor(IndexedColors.BLUE.index);
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFont(headerFont);
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerStyle.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
			Row headerRow = sh.createRow(0);
			for(int i=0; i<columnHeadings.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columnHeadings[i]);
				cell.setCellStyle(headerStyle);
			}
			FileOutputStream fileOut = new FileOutputStream("");
			workbook.write(fileOut);
			workbook.close();
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
