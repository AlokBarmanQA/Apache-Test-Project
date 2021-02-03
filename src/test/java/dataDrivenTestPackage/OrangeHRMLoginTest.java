/**
 * 
 */
package dataDrivenTestPackage;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import dataDrivenUtilities.ExcelDataProviderUtiliy;

/**
 * @author alokb
 *
 */
public class OrangeHRMLoginTest {

	WebDriver driver;
	
	@BeforeMethod
	public void setup() {
		System.setProperty("webdriver.chrome.driver", "./Drivers/chromedriver.exe");
		driver = new ChromeDriver();
		driver.get("https://opensource-demo.orangehrmlive.com/");
	}
	
	@Test(dataProvider="OrangeHRMLoginData", dataProviderClass=ExcelDataProviderUtiliy.class)
	public void orangeHRMLoginTest(String username, String password) throws InterruptedException {
		driver.findElement(By.id("txtUsername")).sendKeys(username);
		driver.findElement(By.id("txtPassword")).sendKeys(password);
		Thread.sleep(3000);
	}
	
	@AfterMethod
	public void tearDoen() {
		driver.quit();
	}
}
