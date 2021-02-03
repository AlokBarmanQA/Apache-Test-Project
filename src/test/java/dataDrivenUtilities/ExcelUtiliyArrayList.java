package dataDrivenUtilities;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtiliyArrayList {
	FileInputStream fis;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	public ExcelUtiliyArrayList(String excelPath) {
		try {
			File source = new File(excelPath);
			fis = new FileInputStream(source);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheet("Sheet1");
			workbook.close();
		} 
		catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	public ArrayList<ExcelUtilityGetterSetter> readExcelData() {
		DataFormatter formatter = new DataFormatter();
		ArrayList <ExcelUtilityGetterSetter> credential = new ArrayList<ExcelUtilityGetterSetter>();
		for(int i=sheet.getFirstRowNum()-1; i<sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			ExcelUtilityGetterSetter object = new ExcelUtilityGetterSetter();
			for(int j=row.getFirstCellNum(); j<row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);
				
				if(j==1) {
					object.setUsername(formatter.formatCellValue(cell));
				}
				if(j==2) {
					object.setPassword(formatter.formatCellValue(cell));
				}
			}
			credential.add(object);
			
		}
		return credential;
	}
}
