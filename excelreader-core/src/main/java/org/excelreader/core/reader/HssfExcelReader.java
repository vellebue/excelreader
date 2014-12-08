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
		super(excelStream);
		Workbook workbook = new HSSFWorkbook(excelStream);
		excelStream.close();
		setDocument(buildDocument(workbook));
	}
	
	private ExcelDocument buildDocument(Workbook workbook) {
		ExcelDocument document = new ExcelDocument();
		for (int i = 0 ; i < workbook.getNumberOfSheets() ; i++) {
			Sheet sheet = workbook.getSheetAt(i);
			String [][] documentSheet = buildDocumentSheet(sheet);
			document.putSheet(sheet.getSheetName(), documentSheet);
		}
		return document;
	}
	
	private String [][] buildDocumentSheet(Sheet sheet) {
		int lastRowNum = sheet.getLastRowNum();
		// Determine the maximum column number
		int lastCellNum = -1;
		for (int i = 0 ; i <= lastRowNum ; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				int currentRowGetLastCellNum = row.getLastCellNum();
				if (currentRowGetLastCellNum - 1 > lastCellNum) {
					lastCellNum = currentRowGetLastCellNum - 1;
				}
			}
		}
		// Transfer data
		String [][] result = new String[lastRowNum + 1][lastCellNum + 1];
		for (int i = 0 ; i <= lastRowNum ; i++) {
			Row row = sheet.getRow(i);
			if (row != null) {
				for (int j = 0 ; j < row.getLastCellNum() ; j++) {
					Cell cell = row.getCell(j);
					if (cell != null) {
						result[i][j] = cell.getStringCellValue();
					}
				}
			}
		}
		return result;
	}

}
