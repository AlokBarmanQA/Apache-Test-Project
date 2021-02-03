package readWriteExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData2 {

	public static void main(String[] args) {
		int rowCount;
		int colCount;
		int i, j;
		String Scenarios;
		String FUND=null;
		String ACCOUNT=null;
		String EXTERNAL_NBR;
		String AMOUNT=null;
		String WAIVE_ALL_FEES;
		String WV_WIRE_FEES;
		String TRADE_DATE;
		String AO_REASON;
		String SP_HAND;
		String DIST_MODE;
		try {
			File file = new File("D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\VastExcelReadWriteData.xlsx");
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			rowCount = sheet.getLastRowNum();
			colCount = sheet.getRow(0).getLastCellNum();
			DataFormatter formatter = new DataFormatter();

			FileOutputStream fos = new FileOutputStream("C:/Users/alokb/OneDrive/Desktop/executionResult.xlsx");
			XSSFWorkbook writeWorkBook = new XSSFWorkbook();
			XSSFSheet writeSheet = writeWorkBook.createSheet("Result");

			for( i=0; i<rowCount; i++) {
				XSSFRow row = sheet.getRow(i);
				Scenarios = formatter.formatCellValue(row.getCell(0));
				FUND = formatter.formatCellValue(row.getCell(1));
				ACCOUNT = formatter.formatCellValue(row.getCell(2));
				EXTERNAL_NBR = formatter.formatCellValue(row.getCell(3));
				AMOUNT = formatter.formatCellValue(row.getCell(4));
				WAIVE_ALL_FEES = formatter.formatCellValue(row.getCell(5));
				WV_WIRE_FEES = formatter.formatCellValue(row.getCell(6));
				TRADE_DATE = formatter.formatCellValue(row.getCell(7));
				AO_REASON = formatter.formatCellValue(row.getCell(8));
				SP_HAND = formatter.formatCellValue(row.getCell(9));
				DIST_MODE = formatter.formatCellValue(row.getCell(10));

				System.out.println(Scenarios+FUND+ACCOUNT+EXTERNAL_NBR+AMOUNT+WAIVE_ALL_FEES+WV_WIRE_FEES+TRADE_DATE+AO_REASON+SP_HAND+DIST_MODE);
				//XSSFRow row = writeSheet.createRow(i); 
				for(j=0; j<5; j++) { 
					row.createCell(j).setCellValue(Scenarios);
					row.createCell(j).setCellValue(FUND);
					row.createCell(j).setCellValue(ACCOUNT);
					row.createCell(j).setCellValue(AMOUNT);
					row.createCell(j).setCellValue("Pass");
				}
				//File writeExcelFile = new File("C:/Users/alokb/OneDrive/Desktop/executionResult.xlsx");
				//FileOutputStream fos = new FileOutputStream(writeExcelFile); XSSFWorkbook
				//writeWorkBook = new XSSFWorkbook(); XSSFSheet writeSheet =
				//writeWorkBook.createSheet("Result");	
				//for( i=0; i<rowCount; i++ ) { 
				//XSSFRow row = writeSheet.createRow(i); 
				//for(j=0; j<5; j++) { 
				//row.createCell(j).setCellValue(Scenarios);
				//row.createCell(j).setCellValue(FUND);
				//row.createCell(j).setCellValue(ACCOUNT);
				//row.createCell(j).setCellValue(AMOUNT);
				//row.createCell(j).setCellValue("Pass");		  
				//}
			}
			writeWorkBook.write(fos);
			workbook.close();
			fos.flush();
			fos.close();
			System.out.println("Written Data in Excel File is completed");
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
}
