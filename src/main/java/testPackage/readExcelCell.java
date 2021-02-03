package testPackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readExcelCell {

	public static void main(String[] args) throws IOException {
		File src=new File("D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\EmployeeData.xlsx.xlsx");
		FileInputStream fis=new FileInputStream(src);
		XSSFWorkbook wb=new XSSFWorkbook(fis);
		XSSFSheet sheet1=wb.getSheetAt(0);
		int rowcount=sheet1.getLastRowNum();
//		int colCount = sheet1.getRow(0).getLastCellNum();
//		 if(MytempCell.getCellType() == Cell.CELL_TYPE_NUMERIC)
//			 int data1 = (int)sheet1.getRow(3).getCell(0).getNumericCellValue();
//		     else if(MytempCell.getCellType() == Cell.CELL_TYPE_STRING)
//		    	 String data0 = sheet1.getRow(1).getCell(0).getStringCellValue();
		System.out.println("Total Row: " + rowcount);

//		String data0 = sheet1.getRow(1).getCell(0).getStringCellValue();
//		System.out.println("Test Data From Excel : "+data0);
		int data1 = (int)sheet1.getRow(3).getCell(0).getNumericCellValue();
		System.out.println("Test Data From Excel : "+data1);

		
		 for(int i=0;i<rowcount+1;i++) { 
		 String data0 = sheet1.getRow(i).getCell(0).getStringCellValue(); 
//		 double data1 = sheet1.getRow(i).getCell(0).getNumericCellValue(); 
		 
		 System.out.println("Test Data From Excel : "+data0);
		 System.out.println("Test Data From Excel : "+data1); }
		 
		wb.close();
	}

}
