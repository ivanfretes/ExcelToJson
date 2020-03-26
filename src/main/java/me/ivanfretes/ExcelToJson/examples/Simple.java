package me.ivanfretes.ExcelToJson.examples;

import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import me.ivanfretes.ExcelToJson.ExcelToJson;

public class Simple {
	ExcelToJson convert;

	public Simple() throws IOException, InvalidFormatException {
		convert = new ExcelToJson("/home/ivan/eclipse-workspace/ExcelToJson/src/main/java/me/ivanfretes/ExcelToJson/examples/file.xlsx");

		
		String[] keyName = { "title1", "title3", "title4" };
		
		convert.setKeyJsonName(keyName);
		convert.setFileJsonName("outputs/file-test.json");
		convert.getSheet(0);
		convert.createFileJSON();
	}

	public static void main(String[] args) throws IOException, InvalidFormatException {
		// TODO Auto-generated method stub

		new Simple();
		
	}

}
