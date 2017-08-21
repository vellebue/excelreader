package org.excelreader.core.reader;

import java.io.InputStream;

import org.junit.Test;

public class XssfExcelReaderTest extends AbstractExcelReaderTest {

	@Override
	public ExcelReader getReaderFromResource(String resourceName)
			throws Exception {
		InputStream stream = getInputStreamFromResource(resourceName);
		return new XssfExcelReader(stream);
	}
	
	@Test
	public void shouldExtractExcelDataFromASingleSheetExcel()
			throws Exception {
		ExcelReader reader = getReaderFromResource("org/excelreader/core/reader/simpleExcel.xlsx");
		super.shouldExtractExcelDataFromASingleSheetExcel(reader);
	}
	
	@Test
	public void shouldExtractExcelDataFromASingleSheetExcelWithSeveralCellTypes()
			throws Exception {
		ExcelReader reader = getReaderFromResource("org/excelreader/core/reader/excelWithSeveralCellTypes.xlsx");
		super.shouldExtractExcelDataFromASingleSheetExcelWithSeveralCellTypes(reader);
	}
}