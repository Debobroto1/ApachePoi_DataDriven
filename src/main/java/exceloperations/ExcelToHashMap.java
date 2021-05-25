package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToHashMap {

	public static void main(String[] args) throws IOException {
		String filepath=".\\Data\\EmpFile.xlsx";
		FileInputStream fis=new FileInputStream(filepath);
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheetAt(0);

		int rows =sheet.getLastRowNum();
		int cell=sheet.getRow(0).getLastCellNum();

		HashMap<String, String> data=new HashMap<String, String>();
		//Reading Data
		for(int r=0;r<rows;r++) {
			String key=sheet.getRow(r).getCell(0).getStringCellValue();
			String value=sheet.getRow(r).getCell(1).getStringCellValue();
			data.put(key,value);
		}
		//Read Data
		for(Map.Entry<String, String> entry:data.entrySet()) {
			System.out.println(entry.getKey()+"|"+entry.getValue());			
		}
	}
}
