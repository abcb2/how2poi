package kj.poi;

import static org.junit.Assert.*;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

public class QuickGuide05 {
	/*
	 * Working with different type of cells
	 * http://poi.apache.org/spreadsheet/quick-guide.html#CellTypes
	 */
	@Test
	public void test() {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");
		Row row = sheet.createRow(2);
		row.createCell(0).setCellValue(1.1);
		row.createCell(1).setCellValue(new Date());
		row.createCell(2).setCellValue(Calendar.getInstance());
		row.createCell(3).setCellValue("a string");
		row.createCell(4).setCellValue(true);
		row.createCell(5).setCellValue(Cell.CELL_TYPE_ERROR);
		
		try {
			FileOutputStream os = new FileOutputStream("workbook.xls");
			wb.write(os);
			os.close();
		} catch (Exception e) {
			fail("fail");
		}
		assertTrue(true);
	}

}
