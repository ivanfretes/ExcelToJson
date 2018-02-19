*ExcelToJson - Parsing the Excel(XLSX) file to JSON format*

***Implement***

```

	public App() throws InvalidFormatException, FileNotFoundException{
		xlsxParseJson = new ExcelToJsonXLSX("~./file.xlsx");
		xlsxParseJson.setInitGrid(0, 0);
		
		// cells ignorates
		String[] ignorate = {"coinciden1","coinciden2"};
		xlsxParseJson.setCellIgnorate(ignorate);
		
		// New keys for the JSON Object
		String[] keyJSONname = {"keyName1","keyName2" , "keyName3"};
		xlsxParseJson.setKeyJsonName(keyJSONname);
		

		// get the sheet by one index
		xlsxParseJson.getSheet(0);
	
		// get the all sheet 
		// xlsxParseJson.getAllSheet();
	
	}

```