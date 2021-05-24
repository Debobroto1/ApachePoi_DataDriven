package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApplyingFormula {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Data");
		XSSFRow row=sheet.createRow(0);
		row.createCell(0).setCellValue(100);
		row.createCell(1).setCellValue(200);
		row.createCell(2).setCellValue(300);
		row.createCell(3).setCellFormula("A1*B1*C1");
		
		FileOutputStream fos=new FileOutputStream(".\\Data\\ProductFile.xlsx");
		workbook.write(fos);
		System.out.println("ProductFile.xlsx created inside Data Folder");
		workbook.close();
		//for existing file
		
	}

}
