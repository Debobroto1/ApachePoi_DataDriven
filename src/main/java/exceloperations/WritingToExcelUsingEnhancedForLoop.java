package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingToExcelUsingEnhancedForLoop {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Sheet 1");

		//Input data
		Object empData[][]= { 
				{"ID" ,"Name","Job"} ,
				{100,"John","Analyst"},
				{101,"Smith","Lead"},
				{102,"Jane","BA"},
				{103,"Doe","L1"},
				{104,"Gayle","PM"},
				{105,"Kyle","Manager"}
		};
		//Total number of rows in input data and cols
		int rows= empData.length;
		int cols= empData[0].length;

		int rownum=0;
		for(Object[] emp:empData) {
			XSSFRow row=sheet.createRow(rownum++);
			int colCount=0;
			for(Object value:emp) {
				XSSFCell cell=row.createCell(colCount);

				if (value instanceof String){
					cell.setCellValue((String)value);
				}
				else if(value instanceof Integer) {
					cell.setCellValue((Integer)value);
				}
				else if(value instanceof Boolean) {
					cell.setCellValue((Boolean)value);
				}

				String filepath=".\\Data\\outputfile.xlsx";
				FileOutputStream output=new FileOutputStream(filepath);
				workbook.write(output);
				output.close();
			}
		}
		System.out.println("File Created!!");
		workbook.close();
	}
}


