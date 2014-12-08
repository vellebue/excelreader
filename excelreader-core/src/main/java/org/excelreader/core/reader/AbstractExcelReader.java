package org.excelreader.core.reader;

import java.io.IOException;
import java.io.InputStream;

import org.excelreader.core.model.ExcelDocument;

public abstract class AbstractExcelReader implements ExcelReader{
	
	private ExcelDocument document;
	
	public AbstractExcelReader(InputStream excelStream) throws IOException {
	}

	public ExcelDocument getDocument() {
		return document;
	}

	protected void setDocument(ExcelDocument document) {
		this.document = document;
	}
}
