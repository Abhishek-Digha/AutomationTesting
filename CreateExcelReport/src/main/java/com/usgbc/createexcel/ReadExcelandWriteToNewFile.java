package com.usgbc.createexcel;

import java.io.IOException;
import com.usgbc.createexcel.CSVToExcelConverter;
import com.usgbc.createexcel.XlsReader;
public class ReadExcelandWriteToNewFile {

	public static void main(String[] args) throws IOException {
		
		
		
		String path ="/var/lib/jenkins/workspace/USGBCSOAPPerformance/AutomationTesting/USGBCSOAPPerformance/Reportfile/SummaryReport.csv";
	    String excelGenerated ="/var/lib/jenkins/workspace/USGBCSOAPPerformance/AutomationTesting/USGBCSOAPPerformance/Reportfile/SummaryReport.xlsx";
	    String ExcelReport= "/var/lib/jenkins/workspace/USGBCSOAPPerformance/AutomationTesting/USGBCSOAPPerformance/Reportfile/UsgbcPerformance.xlsx";
		
		/*String path ="D:\\SummaryReport.csv";
	    String excelGenerated ="D:\\SummaryReport.xlsx";
	    String ExcelReport= "D:\\ExcelReport.xlsx";*/
		
		
		try {
	    CSVToExcelConverter.csvToXLSX(path, excelGenerated);
	    CSVToExcelConverter.CreateExcelreport(ExcelReport);
	    XlsReader oldExcelData= new XlsReader(excelGenerated);
	    XlsReader newExceldata= new XlsReader(ExcelReport);
	    for(int count=2 ; count<= Integer.parseInt(args[0]);count++) {
	    System.out.println(oldExcelData.getCellData("SummaryReport", "label", count));	
	    newExceldata.setCellData("SummaryReport", "SOAP Request", count ,oldExcelData.getCellData("SummaryReport", "label", count));
	    Double time = Double.parseDouble(oldExcelData.getCellData("SummaryReport", "elapsed", count))/1000;
	    newExceldata.setCellData("SummaryReport", "Response Time(In Seconds)", count , time.toString() );
		   
	    }
	
	    
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		
		
		
		
		
	}

}
