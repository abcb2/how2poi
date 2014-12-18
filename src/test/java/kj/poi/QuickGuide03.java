package kj.poi;

import static org.junit.Assert.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

public class QuickGuide03 {
	/*
	 * How to create cells
	 * http://poi.apache.org/spreadsheet/quick-guide.html#CreateCells
	 */
	@Test
	public void test() {
		Workbook wb = new HSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("new sheet");
		
		Row row = sheet.createRow((short)0);
		
		Cell cell = row.createCell(0);
		cell.setCellValue(100);
		row.createCell(1).setCellValue(1.2);
		row.createCell(2).setCellValue(
				createHelper.createRichTextString("This is a ストリング"));
		row.createCell(3).setCellValue(true);
		
		try {
			FileOutputStream os = new FileOutputStream("workbook.xls");
			wb.write(os);
			os.close();
		} catch (Exception e) {
			fail("fail!!");
		}
		assertTrue(true);
	}
}
