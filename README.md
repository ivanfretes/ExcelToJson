#ExcelToJson - Parsing the Excel(XLSX) file to JSON format#

##Implement##

```
	// i.e
	public TestApp() throws InvalidFormatException, IOException{
		xlsxParseJson = new ExcelToJsonXLSX("~/document.xlsx");
		xlsxParseJson.setInitGrid(0, 0);
		
		// cells ignorates
		String[] ignorate = {"dpto.","locales","elect","mesas","telefono", "totales", "zona", "distrito", "condicion"};
		xlsxParseJson.setCellIgnorate(ignorate);
		
		// New keys for the JSON Object
		String[] keyJSONname = {"province" ,"zone" , "name", "elector_cant", "table_cant","phone"};
		xlsxParseJson.setKeyJsonName(keyJSONname);
		

		// get the sheet by one index
		xlsxParseJson.getSheet(0);
		
		// Setting the output file name 
		xlsxParseJson.setFileJsonName("file-name.json");
		xlsxParseJson.createFileJSON();

	}

```