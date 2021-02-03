package employeeDataDriven;

public class EmployeeDataDrivenTest {
	
	public static void main(String[] args) {
		
		EmployeeExcelToObject excelColumn = new EmployeeExcelToObject("D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\EmployeeData.xlsx");
		excelColumn.readEmployeeData();
		
	}

}
