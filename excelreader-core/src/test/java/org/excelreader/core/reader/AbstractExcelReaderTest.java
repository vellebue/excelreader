package org.excelreader.core.reader;

import static org.junit.Assert.*;

import java.io.InputStream;

import org.excelreader.core.model.ExcelDocument;

public abstract class AbstractExcelReaderTest {
	
	protected InputStream getInputStreamFromResource(String resourceName) {
		ClassLoader loader = Thread.currentThread().getContextClassLoader();
		return loader.getResourceAsStream(resourceName);
	}
	
	public abstract ExcelReader getReaderFromResource(String resourceName) throws Exception;
	
	
	public void shouldExtractExcelDataFromASingleSheetExcel(ExcelReader reader) throws Exception{
		ExcelDocument document = reader.getDocument();
		assertEquals(1, document.getNumberOfSheets());
		Object [][] sheet = document.getSheetAt(0);
		assertEquals("John", sheet[0][0]);
		assertEquals("Mary", sheet[0][1]);
		assertEquals("Paul", sheet[0][2]);
		assertEquals("Leila", sheet[1][0]);
		assertEquals("Francis", sheet[1][1]);
		assertEquals("Jessica", sheet[1][2]);
	}
	
	public void shouldExtractExcelDataFromASingleSheetExcelWithSeveralCellTypes(ExcelReader reader) {
		ExcelDocument document = reader.getDocument();
		assertEquals(1, document.getNumberOfSheets());
		Object [][] sheet = document.getSheetAt(0);
		assertEquals("John Doe", sheet[0][0]);
		assertEquals(new Double(1432.26), sheet[0][1]);
		assertEquals(new Double(42208.0), sheet[0][2]);
		assertEquals(Boolean.TRUE, sheet[0][3]);
	}

}
