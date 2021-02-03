package employeeDataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmployeeExcelToObject3 {
	FileInputStream fis;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	Row row;
	int i,j;
	DataFormatter formatter;
	ArrayList<EmployeeGetterSetter>employeeList;
	EmployeeGetterSetter object;
	public EmployeeExcelToObject3(String excelPath) {
		try {
			File source = new File(excelPath);
			fis = new FileInputStream(source);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheet("Sheet1");
			formatter = new DataFormatter();
			employeeList = new ArrayList<EmployeeGetterSetter>();
			for(i=sheet.getFirstRowNum(); i<=sheet.getLastRowNum(); i++) {
				row = sheet.getRow(i);
				object = new EmployeeGetterSetter();
//				for(j=row.getFirstCellNum(); j<row.getLastCellNum(); j++) {
//					Cell cell = row.getCell(j);
//					object.setId(formatter.formatCellValue(row.getCell(0)));
//					object.setFirstName(formatter.formatCellValue(row.getCell(1)));
//					object.setLastName(formatter.formatCellValue(row.getCell(2)));
//					object.setDob(formatter.formatCellValue(row.getCell(3)));

//			}
				employeeList.add(setExcelValue());
				workbook.close();
			}
		} 
		catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	public EmployeeGetterSetter setExcelValue() {
		
		object.setId(formatter.formatCellValue(row.getCell(0)));
		object.setFirstName(formatter.formatCellValue(row.getCell(1)));
		object.setLastName(formatter.formatCellValue(row.getCell(2)));
		object.setDob(formatter.formatCellValue(row.getCell(3)));
		return object;
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
	
	public static void main(String[] args) {
		EmployeeExcelToObject3 excel = new EmployeeExcelToObject3("D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\EmployeeData.xlsx");
		excel.readEmployeeData();
	}

}
