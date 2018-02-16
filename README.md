**ExcelToJson - Parsing the Excel(XLSX) file to JSON format**

***Implement***

```

	public App() throws InvalidFormatException, FileNotFoundException{
		xlsxParseJson = new ExcelToJsonXLSX("~./file.xlsx");
		xlsxParseJson.setInitGrid(0, 0);
		
		// cells ignorates
		String[] ignorate = {"head1","head2"};
		xlsxParseJson.setCellIgnorate(ignorate);
		
		// New keys for the JSON Object
		String[] keyJSONname = {"province","zone" , "name", "elector_cant", "table_cant","phone"};
		xlsxParseJson.setKeyJsonName(keyJSONname);
		

		// get the sheet by one index
		xlsxParseJson.getSheet(0);
	
	}

```