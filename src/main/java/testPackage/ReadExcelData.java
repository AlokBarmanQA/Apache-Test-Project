package testPackage;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelData {

	public static void main(String[] args) throws Exception {

		File file = new File("D:/Eclipse-Workspace/ApacheTestProject/TestData/VastExcelReadWriteData.xlsx");
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(fis);
		XSSFSheet sheet = workbook.getSheet("Sheet1");
		int rowCount = sheet.getLastRowNum();
		int colCount = sheet.getRow(0).getLastCellNum();
		workbook.close();
		for(int i=0; i<rowCount; i++) {
			XSSFRow row = sheet.getRow(i);
			
			for(int j=0; j<colCount; j++) {
				String cellValue = row.getCell(j).toString();
				System.out.print(cellValue +" | ");
			}
			System.out.println();
		}
	}
}
