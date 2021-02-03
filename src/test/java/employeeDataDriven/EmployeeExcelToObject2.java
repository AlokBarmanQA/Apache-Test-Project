package employeeDataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmployeeExcelToObject2 {
	public FileInputStream fis;
	public XSSFWorkbook workbook;
	public XSSFSheet sheet;
	public ArrayList<EmployeeGetterSetter> employeeList;
	public int i, j;

	public EmployeeExcelToObject2() {

	}

	public EmployeeExcelToObject2(String excelPath) {
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

	public ArrayList<EmployeeGetterSetter> setExcelValue() {
		DataFormatter formatter = new DataFormatter();
		employeeList = new ArrayList<EmployeeGetterSetter>();
		for(i=sheet.getFirstRowNum(); i<=sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			EmployeeGetterSetter object = new EmployeeGetterSetter();
			//			for(j=row.getFirstCellNum(); j<row.getLastCellNum(); j++) {
			Cell cell = row.getCell(j);
			object.setId(formatter.formatCellValue(row.getCell(0)));
			object.setFirstName(formatter.formatCellValue(row.getCell(1)));
			object.setLastName(formatter.formatCellValue(row.getCell(2)));
			object.setDob(formatter.formatCellValue(row.getCell(3))); 

			//			}
			employeeList.add(object);
		}
		return employeeList;
	}


	/*
	 * public void readEmployeeData() { EmployeeExcelToObject2 excelColumn = new
	 * EmployeeExcelToObject2(
	 * "D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\EmployeeData.xlsx");
	 * ArrayList<EmployeeGetterSetter>excelData = excelColumn.setExcelValue();
	 * for(EmployeeGetterSetter emp:employeeList) { emp.getId(); emp.getFirstName();
	 * emp.getLastName(); emp.getDob();
	 * 
	 * System.out.println(emp.getId()+" | "+emp.getFirstName()+" | "+emp.getLastName
	 * ()+" | "+emp.getDob()); } }
	 */

	public static void main(String[] args) {
		EmployeeExcelToObject2 excelColumn = new EmployeeExcelToObject2("D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\EmployeeData.xlsx");
//		ArrayList<EmployeeGetterSetter>excelData = excelColumn.setExcelValue();
		for(EmployeeGetterSetter emp:excelColumn.setExcelValue()) {
			emp.getId(); 
			emp.getFirstName(); 
			emp.getLastName(); 
			emp.getDob();

			System.out.println(emp.getId()+" | "+emp.getFirstName()+" | "+emp.getLastName()+" | "+emp.getDob()); } 
	
	}

}
