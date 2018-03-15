package generic.parse.excel.to.json;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class TestApp {
	ExcelToJsonXLSX xlsxParseJson;
	
	public TestApp() throws InvalidFormatException, IOException{
		xlsxParseJson = new ExcelToJsonXLSX("/home/ivan/Documents/PPQ/Datos/migracion/document.xlsx");
		xlsxParseJson.setInitGrid(0, 0);
		
		// cells ignorates
		String[] ignorate = {"dpto.","locales","elect","mesas","telefono", "totales", "zona", "distrito", "condicion"};
		xlsxParseJson.setCellIgnorate(ignorate);
		
		// New keys for the JSON Object
		String[] keyJSONname = {"province" ,"zone" , "name", "elector_cant", "table_cant","phone"};
		xlsxParseJson.setKeyJsonName(keyJSONname);
		

		// get the sheet by one index
		xlsxParseJson.getSheet(0);
		
		xlsxParseJson.setFileJsonName("capital-00.json");
		xlsxParseJson.createFileJSON();
		// get the all sheet 
		// xlsxParseJson.getAllSheet();
			
	
	}
	
    public static void main( String[] args ) throws InvalidFormatException, IOException{
    	new TestApp();
    }
}

