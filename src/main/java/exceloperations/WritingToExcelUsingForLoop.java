package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingToExcelUsingForLoop {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
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

		for(int r=0;r<rows;r++) {
			XSSFRow row=sheet.createRow(r);

			for(int c=0;c<cols;c++) {
				XSSFCell cell=row.createCell(c);

				Object  value=empData[r][c];

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
