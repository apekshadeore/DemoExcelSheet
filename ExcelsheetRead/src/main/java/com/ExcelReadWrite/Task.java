package com.ExcelReadWrite;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;

public class Task {

	public static String result = null;

	public static void main(String[] args) throws Exception {

		System.setProperty("Webdriver.Chrome.driver", "chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("file:///C:/Apeksha%20All%20Softwere/javabykiran-Selenium-Softwares/Offline%20Website/index.html");

		DataFormatter df = new DataFormatter();
		FileInputStream fis = new FileInputStream("Book1.xlsx");
		

		Workbook wb = WorkbookFactory.create(fis);
		Sheet sh = wb.getSheet("Sheet1");
		Cell email, password, c1 = null;

		int rows = sh.getLastRowNum();
		for (int i = 1; i <= rows; i++) {
			email = sh.getRow(i).getCell(0);
			password = sh.getRow(i).getCell(1);
			System.out.println("Username " + email);
			System.out.println("password " + password);

			// clear data
			driver.findElement(By.id("email")).clear();
			driver.findElement(By.id("password")).clear();

			// fullfilldata
			driver.findElement(By.id("email")).sendKeys(df.formatCellValue(email));
			driver.findElement(By.id("password")).sendKeys(df.formatCellValue(password));
			driver.findElement(By.xpath("//button[@type='submit']")).click();

			// if logout button only one time should be click
			List<WebElement> buttonlist = driver.findElements(By.xpath("//a[text()='LOGOUT']"));

			if (buttonlist.size() > 0) {
				result = "pass";
				buttonlist.get(0).click();

			} else 			
				result = "Fail";

            c1 = sh.getRow(i).createCell(2);
			c1.setCellValue(result);

		}
		FileOutputStream fos = new FileOutputStream("Book1.xlsx");
		wb.write(fos);
		wb.close();
		fos.close();

	}

}



