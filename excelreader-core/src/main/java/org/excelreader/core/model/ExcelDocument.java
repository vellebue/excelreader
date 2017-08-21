package org.excelreader.core.model;

import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelDocument {
	
	private List<Map.Entry<String, Object [][]>> sheetsList = new ArrayList<Map.Entry<String,Object[][]>>();
	
	public Object [][] getSheetAt(int index) {
		return sheetsList.get(index).getValue();
	}
	
	public Object [][] getSheetByName(String name) {
		for (Map.Entry<String, Object [][]> entry : sheetsList) {
			if (entry.getKey().equals(name)) {
				return entry.getValue();
			}
		}
		return null;
	}

	public void putSheet(String name, Object [][] sheet) {
		Map.Entry<String, Object [][]> entry = new AbstractMap.SimpleImmutableEntry<String, Object [][]>(name, sheet);
		sheetsList.add(entry);
	}
	
	public int getNumberOfSheets() {
		return sheetsList.size();
	}
}
