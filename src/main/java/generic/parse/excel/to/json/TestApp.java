package generic.parse.excel.to.json;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

public class TestApp {
	ExcelToJsonXLSX xlsxParseJson;
	
	public TestApp() throws InvalidFormatException, IOException{
		xlsxParseJson = new ExcelToJsonXLSX("/home/ivan/Documents/PPQ/Datos/test-ppq.xlsx");
		xlsxParseJson.setInitGrid(2, 0);
		
		String[] ignorate = {"dpto.","locales","elect","mesas","telefono", "totales", "zona", "distrito", "condicion"};
		xlsxParseJson.setCellIgnorate(ignorate);
		
		String[] keyJSONname = {"province" ,"zone" , "name", "elector_cant", "table_cant","phone"};
		xlsxParseJson.setKeyJsonName(keyJSONname);
		
		xlsxParseJson.getSheet(0);
		xlsxParseJson.setFileJsonName("test33.json");
		xlsxParseJson.createFileJSON();
	
	}
	
    public static void main( String[] args ) throws InvalidFormatException, IOException{
    	new TestApp();
    }
}
