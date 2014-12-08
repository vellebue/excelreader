package org.excelreader.core.model;

import java.util.AbstractMap;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ExcelDocument {
	
	private List<Map.Entry<String, String [][]>> sheetsList = new ArrayList<Map.Entry<String,String[][]>>();
	
	public String [][] getSheetAt(int index) {
		return sheetsList.get(index).getValue();
	}
	
	public String [][] getSheetByName(String name) {
		for (Map.Entry<String, String [][]> entry : sheetsList) {
			if (entry.getKey().equals(name)) {
				return entry.getValue();
			}
		}
		return null;
	}

	public void putSheet(String name, String [][] sheet) {
		Map.Entry<String, String [][]> entry = new AbstractMap.SimpleImmutableEntry<String, String [][]>(name, sheet);
		sheetsList.add(entry);
	}
	
	public int getNumberOfSheets() {
		return sheetsList.size();
	}
}
