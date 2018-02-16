package com.parse.apache.poi.ExcelToJson;

import java.io.FileNotFoundException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class App {
	ExcelToJsonXLSX xlsxParseJson;
	
	public App() throws InvalidFormatException, FileNotFoundException{
		xlsxParseJson = new ExcelToJsonXLSX("/home/ivan/Documents/PPQ/Datos/test.xlsx");
		xlsxParseJson.setInitGrid(0, 0);
		
		//String[] ignorate = {"DPTO.","locales de","Electores","Mesas","Telefono", "totales", "zona"};
		//xlsxParseJson.setCellIgnorate(ignorate);
		
		String[] keyJSONname = {"province", "name", "elector_cant", "table_cant", "phone"};
		xlsxParseJson.setKeyJsonName(keyJSONname);
		
		xlsxParseJson.getSheet(0);
	
	}
	
    public static void main( String[] args ) throws InvalidFormatException, FileNotFoundException{
    	new App();
    }
}
