package kj.poi;

import static org.junit.Assert.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

public class QuickGuide06 {
	/*
	 * Iterate over rows and cells
	 * http://poi.apache.org/spreadsheet/quick-guide.html#Iterator
	 */
	@Test
	public void test() {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");
		
		Row row = sheet.createRow(0);
		row.createCell(0).setCellValue("pos0_0");
		row.createCell(1).setCellValue("pos0_1");
		row.createCell(2).setCellValue("pos0_2");
		
		row = sheet.createRow(1);
		row.createCell(0).setCellValue("");
		row.createCell(1).setCellValue("pos1_1");
		row.createCell(2).setCellValue("");
		
		row = sheet.createRow(3);
		row.createCell(2).setCellValue("pos3_2");

		
		try {
			FileOutputStream os = new FileOutputStream("workbook.xls");
			wb.write(os);
			os.close();
		} catch (Exception e) {
			fail("fail");
		}
		assertTrue(true);
		
		try {
			Workbook wb_in = new HSSFWorkbook(new FileInputStream("workbook.xls"));
			Sheet sheet_in = wb_in.getSheetAt(0);
			for(Row row_in : sheet){
				for(Cell cell_in : row_in){
					String str = cell_in.getStringCellValue();
					if(str.equals("")){
						str = "BLANK";
					}
					System.out.print(str);
					System.out.print("\t");
				}
				System.out.print("\n");
			}
		} catch (Exception e) {
			fail("fail");
		}
		assertTrue(true);
	}

}
