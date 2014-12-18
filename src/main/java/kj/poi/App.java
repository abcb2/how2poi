package kj.poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;


public class App{
	@SuppressWarnings("resource")
	public static void main( String[] args ) throws InvalidFormatException, IOException{
		Workbook wb = WorkbookFactory.create(new File("original.xls"));
		for(int i=0; i<wb.getNumberOfSheets(); i++){
			Sheet sheet = wb.getSheetAt(i);
			System.out.println("sheetName: " + sheet.getSheetName());
			Row row = sheet.getRow(0);
			if(row != null){
				Cell cell = row.getCell(0);
				if(cell != null){
					System.out.println(cell.getStringCellValue());
				}
			}
		}
	}
}
