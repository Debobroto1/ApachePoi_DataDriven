package ReadExcelFiles;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadFile {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream fis =new FileInputStream("D:\\testfiles\\ReadingFiles.xlsx");//Locate Workbook
	
		XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
		//XSSFSheet mysheet= myWorkBook.getSheetAt(0);
		XSSFSheet mysheet= myWorkBook.getSheet("Sheet1");

		int totalrow= mysheet.getLastRowNum();
		System.out.println(totalrow);
		XSSFRow myrow= mysheet.getRow(1);
		int lastCellNo=myrow.getLastCellNum();
		System.out.println(lastCellNo);
		XSSFCell mycell = myrow.getCell(2); //0,0
		//XSSFCell to String
		System.out.println(mycell);
		
	
	}

}
