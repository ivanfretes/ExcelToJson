## ExcelToJSON 
Conviente un archivo Excel (.xlxs) a JSON

### Implementación

#### Constructores
```java

	// 1.  Implicito - Sin Parametros
	ExcelToJSON convert = new ExcelToJSON();
	
	// 2. Parametro el nombre del archivo Excel
	ExcelToJSON convert = new ExcelToJSON('~/excel.xlsx');

```

### Metodos
-  Indica desde que celda se va a analizar, las posiciones son como las de una matriz **setInitGrid(1, 0);**, representan filas y columnas
 
	```
		convert.setInitGrid(1, 0);
	```
 
-  Ignora las celdas que contengan esas Palabras **setCellIgnorate()** 

	```
		
		String[] ignoreCellWithWords = {
												"words1","words2","words3","words4",
												"words5","words6","words7","wordsN"
										   }; 
		convert.setCellIgnorate()
	```

-  Los nombres de las claves a ser asignada en cada objeto JSON **setKeyJsonName(keyJSONname);**

	```
		String[] keyJSONname = { "key_1", "key_2", "key_3","key_N"};
		convert.setKeyJsonName(keyJSONname);
	````
 
- Retorna los datos de una Hoja, donde N es un entero que indica el numero de la hoja - **getSheet(N);**

	```
		convert.getSheet(N);
	```

- Actualiza el nombre del Archivo JSON a Generar, si no se usa este metodo, se genera por defecto archivo con el nombre de *xlsx.json* - **setFileJsonName('file-name.json');**

	```
		convert.setFileJsonName("./file-custom-name.json");
	```

- Crea el archivo JSON
 
	```
		convert.createFileJSON();
	```


### Examples
- [Simple JSON File](./examples/Simple.java)
- [Analyse from a Row and Column](./examples/AnalyseRowAndColumn.java)
- [Generate a custom file JSON name](./examples/CustomJsonFileName.java)
- [Ignore some words in cells](./examples/IgnoreWords.java)
