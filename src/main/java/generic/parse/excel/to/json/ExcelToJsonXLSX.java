package generic.parse.excel.to.json;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.JsonElement;
import com.google.gson.JsonParser;

/**
 * @author Iván Fretes
 */
public class ExcelToJsonXLSX {
    private XSSFWorkbook wb;
    private String JSONData;
    private Map<String, ArrayList> sheetResults;
    
    // Name of the column to export as key on the JSON file
    protected String[] KeyJsonName = null;

    // Cells that not inserted/ignorate in the result grid
    protected String[] cellIgnorate = null;

    // Range to column iterate 
    protected int rowIndexInit = 0;
    protected int columnIndexInit = 0;

    // Info and data of files (xlsx and json)
    private File fileInput;
    private String fileExcelName; 
    private final String fileExcelFormat = "xlsx";
    private String fileJsonName = null;
    private BufferedWriter fileOutput;
    
    // Cell content type
    protected String[] cellType = {"BLANK","NUMBER","STRING"};
	

    public ExcelToJsonXLSX(String fExcelName) throws InvalidFormatException, IOException {
        this.initialize(fExcelName);
    }


    // Setting the path or file name of 
    public void setFileExcelName(String fExcelName) {
        try {
            if (fExcelName.toLowerCase().indexOf(this.fileExcelFormat) < 0) {
                throw new Exception("Problem the extension file, This library support a .xlxs format file");
            }
            this.fileExcelName = fExcelName;
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * Generate JSON file and path directory
     * @param fJsonName
     */
    public void setFileJsonName(String fJsonName) {
    	if (null != fJsonName)
    		fileJsonName = this.fileInput.getParent()+"/"+fJsonName;
    	else 
    		fileJsonName = this.fileInput.getParent()+"/xlsx.json";

    }

    public ExcelToJsonXLSX() {}

    // initialize the apps component
    public void initialize(String fExcelName) throws InvalidFormatException, IOException { 
    	this.sheetResults = new HashMap<String, ArrayList>();
        this.setFileExcelName(fExcelName);
        this.openFile();
    }

    
    /**
     * Setting the cell to ignorates
     * @param cellValues
     */
    public void setCellIgnorate(String[] cellValues) {
        this.cellIgnorate = cellValues;
    }

    /**
     * Get the one sheet by index
     * @param sheetIndex
     */
    public void getSheet(int sheetIndex) {
        Sheet sheetTmp = this.wb.getSheetAt(sheetIndex);
        this.getAllRowBySheet(sheetTmp);
        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        this.JSONData = gson.toJson(this.sheetResults);
    }
    
    
    // Return the JSONData generate
    public String getJsonData() {
    	return this.JSONData; 
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
     * @throws IOException 
     */
    protected void createFileJSON() throws IOException  {
    	fileOutput = new BufferedWriter(new FileWriter(fileJsonName));
    	fileOutput.write(this.JSONData);
    	fileOutput.close();
    }

   

    /**
     * Open the file and generate the stream
     * @throws InvalidFormatException
     * @throws FileNotFoundException
     */
    private void openFile() throws InvalidFormatException, FileNotFoundException {
        try {
            this.fileInput = new File(this.fileExcelName);
            OPCPackage pkg = OPCPackage.open(this.fileInput);
            wb = new XSSFWorkbook(pkg);
            pkg.close(); 
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Working all row by one sheet, and validate the rowIndex > rowIndexInit
     * Set the sheet name and add the row to sheetmp    
     * @param sheet
     * @return ArrayList
     */
    private void getAllRowBySheet(Sheet sheet) {
        ArrayList<Map> sheetTmp = new ArrayList<Map>();  
        int rowIndex;
        int cellIndex;
        
        for (Row row : sheet) {
            cellIndex = 0;
            rowIndex = row.getRowNum();
            if (this.verifyRowIndexInit(rowIndex)) {
            	sheetTmp.add(this.getAllCellByRow(row, cellIndex));
            }				
        }
        
        sheetResults.put(sheet.getSheetName(), sheetTmp);
    }


    /**
     * Working the all cell by row, return the row
     * @param row
     * @param cellIndex
     * @return Hashtable
     */
    private Map<String, String> getAllCellByRow(Row row, int cellIndex) {
        Map<String, String> rowTmp = new HashMap<String, String>();
        
        for (Cell cell : row) {
            if (verifycolumnIndexInit(cellIndex)) {
                    if (!this.verifyCellDataEqual(cell) && cellIndex < this.KeyJsonName.length) {
                    	rowTmp.put(this.KeyJsonName[cellIndex],cell.toString());
                    }
            }
            cellIndex++;
        }
        return rowTmp;
    }




    /**
     * Verify if the cell data coincidence with they  
     * @param cell
     * @return boolean
     */
    private boolean verifyCellDataEqual(Cell cell) {
        if (null != this.cellIgnorate) {
            for (String cellIgn : this.cellIgnorate) {
                if (cell.toString().trim().toLowerCase().indexOf(cellIgn.toLowerCase()) > -1) {
                    return true;
                }
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
