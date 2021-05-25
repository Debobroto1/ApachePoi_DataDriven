package exceloperations;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadingPasswordProtected {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		String filepath="";
		String password="";
		FileInputStream fis=new FileInputStream(filepath);
		
		//Workbook workbook=WorkbookFactory.create(fis, password); or
		XSSFWorkbook workbook=(XSSFWorkbook) WorkbookFactory.create(fis, password);
	}

}
