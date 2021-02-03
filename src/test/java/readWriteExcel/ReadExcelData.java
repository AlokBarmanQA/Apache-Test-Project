package readWriteExcel;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData {

	public static void main(String[] args) {

		try {
			File file = new File("C:/Users/alokb/OneDrive/Desktop/excelDataForPOST.xlsx");
	//		File file = new File("D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\VastExcelReadWriteData.xlsx");
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			int rowCount = sheet.getLastRowNum();
			int colCount = sheet.getRow(0).getLastCellNum();
			DataFormatter formatter = new DataFormatter();
			for(int i=0; i<rowCount; i++) {
				XSSFRow row = sheet.getRow(i);

				for(int j=0; j<colCount; j++) {
					String cellValue = formatter.formatCellValue(row.getCell(j));
					System.out.print("  "+cellValue);
				}
				System.out.println();
			}
			workbook.close();		
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
}
