package exceloperations;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMapToExcel {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Emp Data");

		Map<String,String> data=new HashMap<String, String>();
		data.put("JobId", "Role");
		data.put("101", "Manual Tester");
		data.put("102", "Functional Tester");
		data.put("103", "Performance Tester");
		data.put("104", "Automation Engineer");
		data.put("105", "SDET");

		int rownum=0;
		for(Map.Entry<String, String> entry:data.entrySet()) {

			XSSFRow row=sheet.createRow(rownum);
			rownum++;

			row.createCell(0).setCellValue((String)entry.getKey());
			row.createCell(1).setCellValue((String)entry.getValue());

		}
		FileOutputStream fos=new FileOutputStream(".\\Data\\EmpFile.xlsx");
		workbook.write(fos);
		System.out.println("The file is created");
		fos.close();
		workbook.close();
	}

}
