package generic.parse.excel.to.json;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class TestApp {
	ExcelToJsonXLSX xlsxParseJson;
	
	public TestApp() throws InvalidFormatException, IOException{
		
		try {
			xlsxParseJson = new ExcelToJsonXLSX("/home/ivan/Documents/PPQ/Datos/migracion/document.xlsx");
			xlsxParseJson.setInitGrid(0, 0);
			
			// cells ignorates
			String[] ignorate = {"dpto.","locales","elect","mesas",
								 "telefono", "totales", "zona", "distrito", 
								 "condicion", "justicia electoral","cant."};
			xlsxParseJson.setCellIgnorate(ignorate);
			
			// New keys for the JSON Object
			String[] keyJSONname = {"province" ,"zone" , "name", "elector_cant", "table_cant","phone"};
			xlsxParseJson.setKeyJsonName(keyJSONname);
			
			// get the sheet by one index
			xlsxParseJson.getSheet(7);
			
			// Setting name
			xlsxParseJson.setFileJsonName("departament/local-format.json");
			
			// Method to create json file
			xlsxParseJson.createFileJSON();

		} catch (Exception e) {
			e.getMessage();
		}
		
			
	}
	
    public static void main( String[] args ) throws InvalidFormatException, IOException{
    	new TestApp();
    }
}

