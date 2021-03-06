package com.jxl.write;

import java.io.File;
import java.io.IOException;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Runner {

//	static {
//		System.setProperty("webdriver.chrome.driver", "./webdriver/chromedriver/chromedriver.exe");
//	}
	
	static WritableWorkbook writeWB;
	@Test
	public void openBrowser() throws BiffException, IOException, RowsExceededException, WriteException {
		
//		WebDriverManager.chromedriver().setup();
//		WebDriver driver = new ChromeDriver();
//		driver.get("https://mvnrepository.com/");
//		
		String fileLocation = "src/test/resources/Book.xls";
		File file = new File(fileLocation);
		Workbook wb = Workbook.getWorkbook(file);
		writeWB = Workbook.createWorkbook(new File(fileLocation), wb);
		WritableSheet sheet = writeWB.getSheet(0);
		Label label;
		
		label = new Label(0,0,"Cell 0,0"); //Here we put the cell location and the data to write
		sheet.addCell(label);
		writeWB.write();
		writeWB.close();
	}
}
