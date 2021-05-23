package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings("unused")
public class ReadingExcelUsingForLoop {

	@SuppressWarnings("incomplete-switch")
	public static void main(String[] args) throws IOException {
		String excelFilePath=".\\Data\\ReadingFiles.xlsx";
		FileInputStream fis=new FileInputStream(excelFilePath);

		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheetAt(0); //or XSSFSheet sheet=workbook.getSheet("Sheet 1");


		//Using For Loop reading data
		int rows=sheet.getLastRowNum();
		int cols=sheet.getRow(1).getLastCellNum();//tells no. of cell in the given row

		for(int r=0;r<rows;r++) 
		{
			XSSFRow row=sheet.getRow(r);

			for(int c=0;c<cols;c++) 
			{
				XSSFCell cell=row.getCell(c);

				switch(cell.getCellType())
				{
				case STRING:System.out.print(cell.getStringCellValue());
				break;
				case NUMERIC:System.out.print(cell.getNumericCellValue());
				break;
				case BOOLEAN:System.out.print(cell.getBooleanCellValue());
				break;
				}
				System.out.print(" | ");
			}
			System.out.println();
		}	
		workbook.close();
	}
}
