package employeeDataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmployeeExcelToObject {

	ArrayList<EmployeeGetterSetter>employeeList;
	
	public EmployeeExcelToObject(String excelPath) {
		try {
			File source = new File(excelPath);
			FileInputStream fis = new FileInputStream(source);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			DataFormatter formatter = new DataFormatter();
			employeeList = new ArrayList<EmployeeGetterSetter>();
			for(int i=sheet.getFirstRowNum(); i<=sheet.getLastRowNum(); i++) {
				Row row = sheet.getRow(i);
				EmployeeGetterSetter object = new EmployeeGetterSetter();
				for(int j=row.getFirstCellNum(); j<row.getLastCellNum(); j++) {
					Cell cell = row.getCell(j);
					if(j==0) {
						object.setId(formatter.formatCellValue(cell));
					}
					if(j==1) {
						object.setFirstName(formatter.formatCellValue(cell));
					}
					if(j==2) {
						object.setLastName(formatter.formatCellValue(cell));
					}
					if(j==3) {
						object.setDob(formatter.formatCellValue(cell));
					}
				}
				employeeList.add(object);
				workbook.close();
			}
		} 
		catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	public void readEmployeeData() {
		for(EmployeeGetterSetter emp:employeeList) {
			emp.getId();
			emp.getFirstName();
			emp.getLastName();
			emp.getDob();
			System.out.println(emp.getId()+" | "+emp.getFirstName()+" | "+emp.getLastName()+" | "+emp.getDob());
		}
		
	}

}
