package kj.poi;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

public class QuickGuide01 {
	/*
	 * How to create a new workbook
	 * http://poi.apache.org/spreadsheet/quick-guide.html#NewWorkbook
	 */

	@Test
	public void createXLS() {
		Workbook workbook = new HSSFWorkbook();
		try {
			File temp = File.createTempFile("sample", ".xls");
			temp.deleteOnExit();
			FileOutputStream os = new FileOutputStream(temp);
			workbook.write(os);
			os.close();
		} catch (IOException e) {
			e.printStackTrace();
			fail("create workbook fail");
		}
		assertTrue(true);
	}
	
	/*
	 * currentに作成される。シートがないため開くとエラーになる。。
	 */
	@Test
	public void createXLSX(){
		Workbook workbook = new HSSFWorkbook();
		try {
			FileOutputStream os = new FileOutputStream("workbook.xlsx");
			workbook.write(os);
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
			fail("create workbook fail");
		}
		assertTrue(true);
	}
}
