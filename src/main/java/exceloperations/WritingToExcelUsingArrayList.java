package exceloperations;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingToExcelUsingArrayList {

	public static void main(String[] args) throws IOException {

		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Sheet 1");

		//Input data
		ArrayList<Object[]> empData=new ArrayList<Object[]> ();
		empData.add(new Object[]{"ID" ,"Name","Job"});
		empData.add(new Object[]{100,"John","Analyst"});
		empData.add(new Object[]{101,"Smith","Lead"});
		empData.add(new Object[]{102,"Jane","BA"});
		empData.add(new Object[]{103,"Doe","L1"});
		empData.add(new Object[]{104,"Gayle","PM"});
		empData.add(new Object[]{105,"Kyle","Manager"});

		//Total number of rows in input data and cols

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

				String filepath=".\\Data\\outputfileAL.xlsx";
				FileOutputStream output=new FileOutputStream(filepath);
				workbook.write(output);
				output.close();
			}
		}
		System.out.println("File Created!!");
		workbook.close();
	}
}


