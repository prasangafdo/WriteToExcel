package com.jxl.write;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import com.jxl.write.util.ExcelSheetManager;

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
//	@Test
	public void openBrowser() throws BiffException, IOException, RowsExceededException, WriteException {
		
//		WebDriverManager.chromedriver().setup();
//		WebDriver driver = new ChromeDriver();
//		driver.get("https://mvnrepository.com/");
//		
		String fileLocation = "src/test/resources/Book1.xls";
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
	
	@Test
	public void writeNamesToExcel() throws BiffException, IOException, RowsExceededException, WriteException {
		
		ArrayList <String> lstNames = new ArrayList<String>();
		lstNames.add("Prasanga");
		lstNames.add("Dulaj");
		lstNames.add("Kaveesh");
		lstNames.add("Lasan");
		lstNames.add("Piyumi");
		lstNames.add("Shermila");
		
		for(String name: lstNames) {
			System.out.println(name);
		}
		
		ExcelSheetManager excel = new ExcelSheetManager();
		excel.writeNameToSheet(0, lstNames);
	}
	
	@Test
	public void writeDepartmentsToExcel() throws BiffException, IOException, RowsExceededException, WriteException {
		
		ArrayList <String> lstDepartments = new ArrayList<String>();
		lstDepartments.add("Corp Merchendicing");
		lstDepartments.add("EAG");
		lstDepartments.add("Corp Merchendicing");
		lstDepartments.add("Commercial");
		lstDepartments.add("Delivery");
		lstDepartments.add("Delivery");
		
		for(String dept: lstDepartments) {
			System.out.println(dept);
		}
		
		ExcelSheetManager excel = new ExcelSheetManager();
		excel.writeDepartmentToSheet(0, lstDepartments);
	}
}
