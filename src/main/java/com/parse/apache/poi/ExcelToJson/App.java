package com.parse.apache.poi.ExcelToJson;

import java.io.FileNotFoundException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class App {
	ExcelToJsonXLSX xlsxParseJson;
	
	public App() throws InvalidFormatException, FileNotFoundException{
		xlsxParseJson = new ExcelToJsonXLSX("/home/ivan/Documents/PPQ/Datos/test-ppq.xlsx");
		xlsxParseJson.setInitGrid(2, 0);
		
		String[] ignorate = {"dpto.","locales","elect","mesas","telefono", "totales", "zona", "distrito", "condicion"};
		xlsxParseJson.setCellIgnorate(ignorate);
		
		String[] keyJSONname = {"province" ,"zone" , "name", "elector_cant", "table_cant"};
		xlsxParseJson.setKeyJsonName(keyJSONname);
		
		xlsxParseJson.getSheet(9);
	
	}
	
    public static void main( String[] args ) throws InvalidFormatException, FileNotFoundException{
    	new App();
    }
}
