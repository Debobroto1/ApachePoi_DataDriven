package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ApplyinFormulaExistingFile {

	public static void main(String[] args) throws IOException {
		FileInputStream fis=new FileInputStream(".\\Data\\outputfile.xlsx");
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheetAt(0);
		XSSFRow row=sheet.getRow(7);
		row.createCell(0).setCellFormula("A2+A3+A4+A5+A6");		
		fis.close();
		FileOutputStream outPut=new FileOutputStream(".\\\\Data\\\\outputfile.xlsx");
		workbook.write(outPut);
		outPut.close();
		workbook.close();
		

	}

}
