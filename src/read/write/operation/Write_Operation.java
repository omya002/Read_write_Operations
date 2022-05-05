package read.write.operation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class Write_Operation {

	public static void main(String[] args) throws IOException, InterruptedException {
		// TODO Auto-generated method stub
		System.setProperty("webdriver.chrome.driver", "D:\\Softwares\\Selenium\\chromedriver.exe");

		File file =    new File("D:\\test123xls.xls");

		FileInputStream inputStream = new FileInputStream(file);

		HSSFWorkbook wb=new HSSFWorkbook(inputStream);

		HSSFSheet sheet = wb.getSheet("test");

		int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();

		WebDriver driver = new ChromeDriver();

		driver.manage().window().maximize();
		driver.get("https://demoqa.com/automation-practice-form");

		//System.out.println("123");

		WebElement firstName=driver.findElement(By.id("firstName"));
		WebElement lastName=driver.findElement(By.id("lastName"));
		WebElement email=driver.findElement(By.id("userEmail"));
		WebElement genderMale= driver.findElement(By.id("gender-radio-1"));
		WebElement mobile=driver.findElement(By.id("userNumber"));
		WebElement address=driver.findElement(By.id("currentAddress"));
		WebElement submitBtn=driver.findElement(By.id("submit"));

		//Thread.sleep(2000);

		for(int i=1;i<=rowCount;i++) {

			Thread.sleep(1000);

			firstName.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
			lastName.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
			email.sendKeys(sheet.getRow(i).getCell(2).getStringCellValue());

			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].click();", genderMale);

			Thread.sleep(1000);

			mobile.sendKeys(sheet.getRow(i).getCell(4).getStringCellValue());
			address.sendKeys(sheet.getRow(i).getCell(5).getStringCellValue());

			//System.out.println("123");

			submitBtn.click();

			WebElement confirmationMessage = driver.findElement(By.xpath("//div[text()='Thanks for submitting the form']"));

			HSSFCell cell = sheet.getRow(i).createCell(6);

			if (confirmationMessage.isDisplayed()) {

				cell.setCellValue("PASS");

			} else {

				cell.setCellValue("FAIL");
			}
			FileOutputStream outputStream = new FileOutputStream("D:\\test123xls.xls");
			wb.write(outputStream);

			WebElement closebtn = driver.findElement(By.id("closeLargeModal"));
			closebtn.click();

			driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));

		}
		wb.close();
		driver.quit();

	}

}

