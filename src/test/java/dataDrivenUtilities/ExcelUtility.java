package dataDrivenUtilities;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {
	FileInputStream fis;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	public ExcelUtility(String excelPath) {
		
		try {
			File source = new File(excelPath);
			fis = new FileInputStream(source);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheet("Sheet1");
		} catch (Exception e) {
			e.printStackTrace();
		}
		
	}
	
	public String getExcelDataOneRow(int row, int column) {
		DataFormatter formatter = new DataFormatter();
		String data = formatter.formatCellValue(sheet.getRow(row).getCell(column));
		System.out.println(data);
		return data;
	}
	
	public String[][] getExcelDataString() {
		
		int rowCount =sheet.getPhysicalNumberOfRows();
		int colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		System.out.println(rowCount+"< --- >"+colCount);
		DataFormatter formatter = new DataFormatter();
		String[][] data = new String[rowCount][colCount];
		for(int i=0; i<rowCount; i++) {
			for(int j=0; j<colCount; j++) {
				String cellData = formatter.formatCellValue(sheet.getRow(i).getCell(j));
				data[i][j]=cellData;
				System.out.print(cellData+" | ");
			}
			System.out.println();
		}
		return data;
	}
	
	public Object[][] getExcelDataObject() {
		int rowCount = sheet.getPhysicalNumberOfRows();
		int colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		DataFormatter formatter = new DataFormatter();
		Object[][] data = new Object[rowCount][colCount];
		for(int i=0; i<rowCount; i++) {
			for(int j=0; j<colCount; j++) {
				String cellData = formatter.formatCellValue(sheet.getRow(i).getCell(j));
				data[i][j] = cellData;
				System.out.print(cellData+" | ");
			}
			System.out.println();
		}
		return data;
	}
	
	
	
	public static void main(String[] args) {
		ExcelUtility excel = new ExcelUtility("D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\OrangeHRMLoginData.xlsx.xlsx");
		excel.getExcelDataOneRow(1, 0);
		excel.getExcelDataString();
		excel.getExcelDataObject();
	}

}
