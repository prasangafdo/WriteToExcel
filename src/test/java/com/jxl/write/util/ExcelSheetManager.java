package com.jxl.write.util;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelSheetManager {
	
	private File file;
	private Workbook wb;
	private WritableWorkbook writeWB;
	private String fileLocation;

	private void getWorkbook() throws BiffException, IOException {
		fileLocation = "src/test/resources/Book1.xls";
		file = new File(fileLocation);
		wb = Workbook.getWorkbook(file);
		
	}
	
//	public String[][] getWorkBookData(){//Will be completed later
//		return null;
//		
//	}
//	
	public void writeNameToSheet(int SheetNumber, ArrayList<String> data) throws BiffException, IOException, RowsExceededException, WriteException {
		this.getWorkbook();
		writeWB = Workbook.createWorkbook(new File(fileLocation), wb);
		WritableSheet sheet = writeWB.getSheet(SheetNumber);
		Label label;
		for(int i=0;i<data.size();i++) {
			label = new Label(0,i,data.get(i)); //Here we put the cell location and the data to write
			sheet.addCell(label);
		}
		
		writeWB.write();
		writeWB.close();
	}
	
	public void writeDepartmentToSheet(int SheetNumber, ArrayList<String> data) throws BiffException, IOException, RowsExceededException, WriteException {
		this.getWorkbook();
		writeWB = Workbook.createWorkbook(new File(fileLocation), wb);
		WritableSheet sheet = writeWB.getSheet(SheetNumber);
		Label label;
		for(int i=0;i<data.size();i++) {
			label = new Label(1,i,data.get(i)); //Here we put the cell location and the data to write
			sheet.addCell(label);
		}
		
		writeWB.write();
		writeWB.close();
	}
}
