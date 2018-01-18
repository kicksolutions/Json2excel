/**
 * 
 */
package com.kicksolutions.Json2Excel.excel;

import java.awt.Color;
import java.io.BufferedReader;
import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jackson.JsonFactory;
import org.codehaus.jackson.JsonParseException;
import org.codehaus.jackson.JsonParser;
import org.codehaus.jackson.JsonToken;
/**
 * @author MBARDIYA
 * 
 */
public class ExcelCodegen {

	private static final Logger LOGGER = Logger.getLogger(ExcelCodegen.class.getName());
	
	private File targetLocation;
	private String jsonFileName;
	private Map<String, String> keyValueMap = new HashMap<String, String>();

	/**
	 * @param targetLocation2 
	 * 
	 */
	public ExcelCodegen(String jsonFileName, File targetLocation) 
	{
		this.jsonFileName = jsonFileName;
		this.targetLocation = targetLocation;
	}

	public List<String> OutputLanguageList(String languagesfile)
	{
		List<String> listOfLanguegs = null;
		try {
    		BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(languagesfile)));
    		String strLine;
    		listOfLanguegs = new ArrayList<>();
    		while ((strLine = br.readLine()) != null)  
    		{
    		  listOfLanguegs.add(strLine);
    		}
    		br.close();
		} 
    	catch (IOException e) 
    	{
				e.printStackTrace();
		}
		return listOfLanguegs;
  }

	/**
	 * 
	 */
	public String generateExcel(String languagesfilePath) throws IOException, IllegalAccessException 
	{
		LOGGER.entering(LOGGER.getName(), "generateExcel");
		List<String> listOfLanguegs = OutputLanguageList(languagesfilePath);
		parseJson();
		
		Writer writer = null;
		String excelFilePath = new StringBuilder().append(targetLocation.getAbsolutePath()).append(File.separator)
				.append("Translation.xls").toString();  //TODO
		try 
		{
			//TODO
			createExcel(listOfLanguegs, excelFilePath);
			LOGGER.log(Level.FINEST, "Sucessfully Written Excel File @ " + excelFilePath);
		} catch (Exception e)
		{
			LOGGER.log(Level.SEVERE, e.getMessage(), e);
			throw new IllegalAccessException(e.getMessage());
		} finally {
			if (writer != null) {
				writer.flush();
			}
		}

		LOGGER.exiting(LOGGER.getName(), "generateExcel");
		return excelFilePath;
	}

	private void createExcel(List<String> listOfLanguegs, String excelFilePath) throws FileNotFoundException, IOException
	{
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Translation Details");
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) 11);
		font.setFontName(XSSFFont.DEFAULT_FONT_NAME);
		font.setBoldweight(XSSFFont.COLOR_NORMAL);
		font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		font.setColor(new XSSFColor(new Color(0, 34, 102)).getIndexed());
		style.setFont(font);
		style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		CellStyle cs = workbook.createCellStyle();
		cs.setWrapText(true);
		
		//create first row with 2 columns
		XSSFRow rowhead = sheet.createRow((short) 0);
		rowhead.setRowStyle((XSSFCellStyle) style);
		createCell(0, style, rowhead, "Key");
		createCell(1, style, rowhead, "ValueFromJson");

		for(int count=0; count < listOfLanguegs.size(); count++)
		{
			createCell(count+2, style, rowhead, listOfLanguegs.get(count).toString());
			//System.out.println("languages are : " + listOfLanguegs.get(count).toString());
		}
		
		//prepare row further
		    int rowCount = 0;
			String key = null, value = null;
			Set mapSet = (Set) keyValueMap.entrySet();
	        Iterator mapIterator = mapSet.iterator();
	        while (mapIterator.hasNext()) 
	        {
	        	rowCount++;
				XSSFRow row = sheet.createRow((short) rowCount);
				row.setRowStyle((XSSFCellStyle) style);
	        	Map.Entry mapEntry = (Map.Entry) mapIterator.next();
	        	key = (String) mapEntry.getKey();
	        	value = (String)mapEntry.getValue();
	        	//System.out.println("Key : " + key + "= Value : " + value);
	        	
	        	createCell(0, style, row, key); // TODO put key from map
				createCell(1, style, row, value); //  TODO put value form map
				
				//TODO - call google api to conevrt
				
				for(int colCount = 2; colCount < listOfLanguegs.size()+2; colCount++)
				{
					createCell(colCount, style, row, value); //TODO put converted value of map value
				}
	        }

		FileOutputStream fileOut = new FileOutputStream(excelFilePath);
		workbook.write(fileOut);
		//workbook.close(); //TODO
	}

	private void createCell(int count, CellStyle style, XSSFRow rowhead, String keyName)
	{
		XSSFCell cell = rowhead.createCell(count);
		cell.setCellStyle(style);
		cell.setCellValue(keyName);
	}
	
	
	private Map<String, String> parseJson() throws JsonParseException, IOException
	  {
		JsonParser jsonParser = new JsonFactory().createJsonParser(new File(jsonFileName));
		jsonParser.nextToken();
		while(jsonParser.nextToken() != JsonToken.END_OBJECT)
		{
			String key = jsonParser.getCurrentName();
			jsonParser.nextToken();
			String value = jsonParser.getText();
			keyValueMap.put(key, value);
			//System.out.println("key is : " + key + "and Value is : " + value);
		}
		return keyValueMap;  
	  }
}