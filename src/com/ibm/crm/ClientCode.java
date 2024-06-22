package com.ibm.crm;

public class ClientCode {
	
	public static void main(String[] args) {	
		try {
			ReadingWritingExcelUtility rweu = new ReadingWritingExcelUtility();
			//rweu.writingExcelData();
			rweu.readingExcelData();
		} catch (Exception e) {
			e.printStackTrace();
		} 
	}

}
