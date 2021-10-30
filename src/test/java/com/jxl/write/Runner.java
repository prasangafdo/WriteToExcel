package com.jxl.write;

import java.io.File;
import java.io.IOException;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableWorkbook;

public class Runner {

//	static {
//		System.setProperty("webdriver.chrome.driver", "./webdriver/chromedriver/chromedriver.exe");
//	}
	
	static WritableWorkbook writeWB;
	@Test
	public void openBrowser() throws BiffException, IOException {
		
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.get("https://mvnrepository.com/");
		
		String fileLocation = "src/test/resources/ExcelToWrite.xls";
		File file = new File(fileLocation);
		Workbook wb = Workbook.getWorkbook(file);
		writeWB = Workbook.createWorkbook(new File(fileLocation), wb);
		
		
		
	}
}
