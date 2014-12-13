package org.excelreader.core.reader;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.excelreader.core.model.ExcelDocument;

public class HssfExcelReader extends AbstractExcelReader { 

	public HssfExcelReader(InputStream excelStream) throws IOException {
		super(new HSSFWorkbook(excelStream));
		//Workbook workbook = new HSSFWorkbook(excelStream);
		excelStream.close();
		//setDocument(buildDocument(workbook));
	}
	
	

}
