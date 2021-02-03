package dataDrivenUtilities;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;

public class ExcelDataProviderUtiliy {

	static FileInputStream fis;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	public static Object[][] getExcelDataObject(String excelPath) {
		try {
			File source = new File(excelPath);
			fis = new FileInputStream(source);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheet("Sheet1");
		} catch (Exception e) {
			e.printStackTrace();
		}
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
	
	@DataProvider(name="OrangeHRMLoginData")
	public static Object[][] dataProviderMethod(){
		Object[][] data = getExcelDataObject("D://Eclipse-Workspace/ApacheTestProject/TestData/OrangeHRMLoginData.xlsx.xlsx");
		return data;
	}
}
