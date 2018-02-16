package com.parse.apache.poi.ExcelToJson;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Objects;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

/**
 * @author Iván Fretes
 */
public class ExcelToJsonXLSX {
	
	
	
	// Excel data / config
	private String fileExcelName;
	private final String fileExcelFormat = "xlsx";
	
	private XSSFWorkbook wb;
	private String JSONData;
	
	// Name of the column to export as key on the JSON file
	protected String[] KeyJsonName = null;
	
	// Cells that not inserted/ignorate in the result grid
	protected String[] cellIgnorate = null;
	
	// Range to column iterate 
	protected int rowIndexInit = 0;
	protected int columnIndexInit = 0;
	

	public ExcelToJsonXLSX(String fExcelName) throws InvalidFormatException, FileNotFoundException {
		this.initialize(fExcelName);
	}
	
	
	public ExcelToJsonXLSX() {}
	
	public void initialize(String fExcelName) throws InvalidFormatException, FileNotFoundException {
		this.fileExcelName = fExcelName; 
		this.openFile();
		this.createFileJSON("/home/ivan/Documents/PPQ/Datos/xlsx.json"); // (improve)
	}
	
	public void setCellIgnorate(String[] cellValues) {
		this.cellIgnorate = cellValues;
	}
	
	/**
	 * Get the one sheet by index
	 * @param sheetIndex
	 */
	public void getSheet(int sheetIndex) {
		Sheet sheetTmp = this.wb.getSheetAt(sheetIndex);
		Gson gson = new GsonBuilder().setPrettyPrinting().create();
		JsonParser jp = new JsonParser();
		JsonElement je = jp.parse(this.getAllRowBySheet(sheetTmp).toString());
		System.out.println(gson.toJson(je));
		
		
		// hacer put en el JSON

	}
	
	/**
	 * Return the number of the cant sheet
	 * @return int
	 */
	public int getSheetNumber() {
		return this.wb.getNumberOfSheets();
	}
	
	/**
	 * Get the all sheet
	 */
	public void getAllSheet() {
		int sheetNumber = this.getSheetNumber();
		for (int i = 0; i < sheetNumber; i++) {
			this.getSheet(i);
		}
	}
	
	/**
	 * Create the  filename.json
	 * @param fileName
	 * @throws FileNotFoundException
	 */
	protected void createFileJSON(String fileName) throws FileNotFoundException  {
		FileOutputStream out = new FileOutputStream(fileName);
		//out.
		
	}
	

	/**
	 * Open the file and generate the stream
	 * @throws InvalidFormatException
	 * @throws FileNotFoundException
	 */
	private void openFile() throws InvalidFormatException, FileNotFoundException {
		try {
			File fileInput = new File(this.fileExcelName);
			OPCPackage pkg = OPCPackage.open(fileInput);
			wb = new XSSFWorkbook(pkg);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Working all row by one sheet, and validate the rowIndex > rowIndexInit  
	 * @param sheet
	 * @return JSONObject
	 */
	private JSONObject getAllRowBySheet(Sheet sheet) {
		JSONArray jsonArray = new JSONArray();
		JSONObject jsonObject = new JSONObject();
		
		int rowIndex;
		int cellIndex;
		for (Row row : sheet) {
			cellIndex = 0;
			rowIndex = row.getRowNum();
			if (this.verifyRowIndexInit(rowIndex)) {
				jsonArray.add(this.getAllCellByRow(row, cellIndex));
			}				
	    }
		jsonObject.put(sheet.getSheetName(), jsonArray);
		return jsonObject;
	}
	
	
	/**
	 * Working the all cell by row, return the row
	 * @param row
	 * @param cellIndex
	 * @return JSONObject
	 */
	private JSONObject getAllCellByRow(Row row, int cellIndex) {
		JSONObject jsonObject = new JSONObject();
		for (Cell cell : row) {
			if (verifycolumnIndexInit(cellIndex)) {
				if (!this.verifyCellDataEqual(cell)) {
					jsonObject.put(this.KeyJsonName[cellIndex], cell.toString().replaceAll("  ","").trim());
				}
			}
			cellIndex++;
		}
		return jsonObject;
	}
	
	
	/**
	 * Verify if the cell data coincidence with they  
	 * @param cell
	 * @return boolean
	 */
	private boolean verifyCellDataEqual(Cell cell) {
		if (null != this.cellIgnorate) {
			for (String cellIgn : this.cellIgnorate) {
				if (cell.toString().toLowerCase().trim().indexOf(cellIgn.toLowerCase()) > -1)
					return true;
			}
		}
		return false;
	}
	
	
	// Verify the rowIndexInit 	
	public boolean verifyRowIndexInit(int rowIndex) {
		return this.rowIndexInit <= rowIndex;
	}
	
	
	// Verify the columnIndexInit 
	public boolean verifycolumnIndexInit(int columnIndex) {
		return this.columnIndexInit <= columnIndex;
	}

	/**
	 * Setting the new key, of parameter data, Getting the columns data()
	 * @param keysName
	 */
	public void setKeyJsonName(String[]  keysName) {
		this.KeyJsonName = keysName;
	}
	
	/**
	 * Set the rowIndexInit & columnIndexInit of the grid or sheet
	 * @param rowInit
	 * @param columnInit
	 */
	public void setInitGrid(int rowInit, int columnInit) {
		this.rowIndexInit = this.naturalNumber(rowInit);
		this.columnIndexInit = this.naturalNumber(columnInit);
	}
	
	// Convert the integer number in a  natural number
	private int naturalNumber(int nmb) {
		return nmb < 0 ? (nmb * -1) : nmb;
	}

}
