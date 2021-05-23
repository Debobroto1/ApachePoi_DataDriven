package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingExcelIteratorExample {

	@SuppressWarnings("incomplete-switch")
	public static void main(String[] args) throws IOException {
		String excelFilePath=".\\Data\\ReadingFiles.xlsx";
		FileInputStream fis=new FileInputStream(excelFilePath);
		XSSFWorkbook workbook=new XSSFWorkbook(fis);
		XSSFSheet sheet=workbook.getSheetAt(0);

		int totalrow=sheet.getLastRowNum();
		int totalcell=sheet.getRow(0).getLastCellNum();

		System.out.println("Total number of row "+totalrow);
		System.out.println("Total number of cell "+totalcell);

		Iterator rowIterator=sheet.iterator();
		while(rowIterator.hasNext()) {
			
			XSSFRow	row =(XSSFRow)rowIterator.next();
			Iterator cellIterator= row.cellIterator();
			
			while(cellIterator.hasNext()) {
				XSSFCell cell=(XSSFCell) cellIterator.next();

				switch(cell.getCellType()) {
				case STRING:System.out.print(cell.getStringCellValue());
				break;
				case NUMERIC:System.out.print(cell.getNumericCellValue());
				break;
				case BOOLEAN:System.out.print(cell.getBooleanCellValue());
				break;
				}
				System.out.print("|");
			}
			System.out.println();
		}
	}

}
