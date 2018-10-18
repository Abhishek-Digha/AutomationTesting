package com.usgbc.createexcel;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class CSVToExcelConverter {

    	public static String sheetName="SummaryReport";
		public static String colNameOne= "SOAP Request";
		public static String colNameTwo= "Response Time(In Seconds)";
		
		public static void csvToXLSX(String csvfile, String excelFile) {
		    try {
		        String csvFileAddress = csvfile; //csv file address
		        String xlsxFileAddress = excelFile; //xlsx file address
		        XSSFWorkbook workBook = new XSSFWorkbook();
		        XSSFSheet sheet = workBook.createSheet(sheetName);
		        String currentLine=null;
		        int RowNum=0;
		        BufferedReader br = new BufferedReader(new FileReader(csvFileAddress));
		        while ((currentLine = br.readLine()) != null) {
		            String str[] = currentLine.split(",");
		           
		            XSSFRow currentRow=sheet.createRow(RowNum);
		            RowNum++;
		            for(int i=0;i<str.length;i++){
		                currentRow.createCell(i).setCellValue(str[i]);
		            }
		        }

		        FileOutputStream fileOutputStream =  new FileOutputStream(xlsxFileAddress);
		        workBook.write(fileOutputStream);
		        fileOutputStream.close();
		        System.out.println("Done");
		    } catch (Exception ex) {
		        System.out.println(ex.getMessage()+"Exception in try");
		    }
		}
    	
    	
    	public static void CreateExcelreport(String fileName) {
    		

    		
    		try {
    			XSSFWorkbook workbook = new XSSFWorkbook();
    			XSSFSheet sheet = workbook.createSheet(sheetName);  

                XSSFRow rowhead = sheet.createRow((short)0);
                rowhead.createCell(0).setCellValue(colNameOne);
                rowhead.createCell(1).setCellValue(colNameTwo);
               
                FileOutputStream fileOut = new FileOutputStream(fileName);
                workbook.write(fileOut);
                fileOut.close();
                System.out.println("Your excel file has been generated!");

            } catch ( Exception ex ) {
                System.out.println(ex);
            }
    		
    	}
}
