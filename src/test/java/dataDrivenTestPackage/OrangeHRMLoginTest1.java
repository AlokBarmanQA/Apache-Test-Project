/**
 * 
 */
package dataDrivenTestPackage;

import java.util.ArrayList;

import org.openqa.selenium.By;
import org.testng.annotations.Test;
import apacheTestBase.BaseClass;
import dataDrivenUtilities.ExcelUtilityGetterSetter;
import dataDrivenUtilities.ExcelUtiliyArrayList;

/**
 * @author alokb
 *
 */
public class OrangeHRMLoginTest1 extends BaseClass{
	
	public OrangeHRMLoginTest1() {
		super();
	}

	
	@Test
	public void orangeHRMLoginTest() throws InterruptedException {
		ExcelUtiliyArrayList excel = new ExcelUtiliyArrayList("D:\\Eclipse-Workspace\\ApacheTestProject\\TestData\\OrangeHRMLoginData.xlsx.xlsx");
		ArrayList<ExcelUtilityGetterSetter> columnData = excel.readExcelData();
		for(ExcelUtilityGetterSetter data : columnData){
		driver.findElement(By.id("txtUsername")).sendKeys(data.getUsername());
		driver.findElement(By.id("txtPassword")).sendKeys(data.getPassword());
		Thread.sleep(3000);
		}
	}

}
