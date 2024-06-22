package com.ibm.crm;

import java.io.File;
import java.io.IOException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ReadingWritingExcelUtility {

	String filePath = "C:\\demo\\output.xls";
	
	public void writingExcelData() throws IOException, RowsExceededException, WriteException {
		File file = new File(filePath);
		WritableWorkbook wb = Workbook.createWorkbook(file);
		WritableSheet sht = wb.createSheet("data", 0);
		Label ll = new Label(0, 0, "India");
		sht.addCell(ll);
		sht.addCell(new Label(0, 1, "Japan"));
		sht.addCell(new Label(0, 2, "Koria"));
		wb.write();
		wb.close();
		System.out.println("Workbook is created");
	}
	
	public void readingExcelData() throws BiffException, IOException {
		File file = new File(filePath);
		Workbook wb = Workbook.getWorkbook(file);
		Sheet sheet = wb.getSheet(0);
		int noOfRows = sheet.getRows();
		int noOfColumns = sheet.getColumns();
		System.out.println("No of Rows: " + noOfRows);
		System.out.println("No of Columns: " + noOfColumns);
		System.out.println("Cell02 1st Column and 3rd Row: " + sheet.getCell(0, 2).getContents());
		String data = "";
		for (int i = 0; i < noOfRows; i++) {
			for (int j = 0; j < noOfColumns; j++) {
				data = sheet.getCell(j, i).getContents();
				System.out.print(data + "\t");
			}
			System.out.println();
		}
		wb.close();
	}
	
}
