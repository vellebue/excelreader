package org.excelreader.core.reader;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class HssfExcelReader extends AbstractExcelReader { 

	public HssfExcelReader(InputStream excelStream) throws IOException {
		super(new HSSFWorkbook(excelStream));
		excelStream.close();
	}
}
