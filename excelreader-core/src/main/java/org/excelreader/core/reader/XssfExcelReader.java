package org.excelreader.core.reader;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XssfExcelReader extends AbstractExcelReader {

	public XssfExcelReader(InputStream excelStream) throws IOException {
		super(new XSSFWorkbook(excelStream));
		excelStream.close();
	}

}
