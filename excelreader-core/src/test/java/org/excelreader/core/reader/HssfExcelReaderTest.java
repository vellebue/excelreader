package org.excelreader.core.reader;

import static org.junit.Assert.*;

import java.io.InputStream;

import org.excelreader.core.model.ExcelDocument;
import org.junit.Test;

public class HssfExcelReaderTest {
	
	@Test
	public void shouldExtractExcelDataFromASingleSheetExcel() throws Exception{
		InputStream stream = Thread.currentThread().getContextClassLoader()
				.getResourceAsStream("org/excelreader/core/reader/simpleExcel.xls");
		HssfExcelReader reader = new HssfExcelReader(stream);
		ExcelDocument document = reader.getDocument();
		assertEquals(1, document.getNumberOfSheets());
		String [][] sheet = document.getSheetAt(0);
		assertEquals("John", sheet[0][0]);
		assertEquals("Mary", sheet[0][1]);
		assertEquals("Paul", sheet[0][2]);
		assertEquals("Leila", sheet[1][0]);
		assertEquals("Francis", sheet[1][1]);
		assertEquals("Jessica", sheet[1][2]);
	}

}
