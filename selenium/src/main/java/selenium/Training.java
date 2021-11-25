package selenium;

import java.io.File;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class Training {

	WebDriver driver;

	@Parameters("browserName")
	@BeforeTest
	public void InitialiseBrowser(String browserName) {
		switch (browserName.toLowerCase()) {
		case "chrome":
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
			break;
		case "edge":
			WebDriverManager.edgedriver().setup();
			driver = new EdgeDriver();
			break;
		case "firefox":
			WebDriverManager.firefoxdriver().setup();
			driver = new FirefoxDriver();
			break;
		default:
			System.err.println("Browser name is Invalid");
			break;
		}

		driver.manage().window().maximize();
	}

	@AfterTest
	public void Teardown() {
		driver.quit();
	}

	@Parameters("url")
	@Test
	public void LaunchApp(String url) {
		driver.get(url);
	}

	/*
	 * @Test public void MainDetails() throws Exception {
	 * 
	 * Thread.sleep(1000); driver.findElement(By.id("maindetailsTabId")).click();
	 * Thread.sleep(1000);
	 * 
	 * }
	 */

	@Test(dataProvider = "test1data")
	public void NewUserCreation(String firstname, String lastname, String address, String mobileno, String emailid)
			throws Exception {
		Thread.sleep(3000);

		driver.findElement(By.id("addNewId")).click();
		Thread.sleep(3000);

		driver.findElement(By.id("firstName")).sendKeys(firstname);
		driver.findElement(By.id("lastName")).sendKeys(lastname);
		driver.findElement(By.id("address")).sendKeys(address);
		driver.findElement(By.id("mobileNo")).sendKeys(mobileno);
		driver.findElement(By.id("email")).sendKeys(emailid);

		Thread.sleep(3000);

		driver.findElement(By.id("submitId")).click();
		Thread.sleep(3000);

	}

	@DataProvider(name = "test1data")
	public Object[][] getData() {
		String excelPath = "D:/TestProject/selenium/excel/data.xlsx";
		File f = new File(this.getClass().getProtectionDomain().getCodeSource().getLocation().getPath());
		System.out.print("ss1:" + f.getAbsolutePath() + "/data.xlsx");
		Object data[][] = testData(excelPath, "Sheet1");
		return data;

	}

	@SuppressWarnings("static-access")
	public static Object[][] testData(String excelPath, String SheetName) {

		ExcelUtils excel = new ExcelUtils(excelPath, SheetName);

		int rowCount = excel.getRowCount();
		int colCount = excel.getColCount();

		Object data[][] = new Object[rowCount - 1][colCount];

		for (int i = 1; i < rowCount; i++) {
			for (int j = 0; j < colCount; j++) {
				String cellData = ExcelUtils.getCellDatastring(i, j);
				data[i - 1][j] = cellData;
			}
		}
		System.out.println();

		return data;
	}

}
