package com.kicksolutions.Json2Excel.excel;

import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.xml.parsers.SAXParser;

/**
 * MBARDIYA
 *
 */
public class ExcelGenerator 
{
	private static final Logger LOGGER = Logger.getLogger(ExcelGenerator.class.getName());
	
	public ExcelGenerator() {
		super();
	}
	    
    /**
     * 
     * @param jsonFile
     * @param output
     */
    public void transformJson2Excel(String jsonFile,String output, String languagesfilePath){
    	LOGGER.entering(LOGGER.getName(), "transformJson2Excel");
    	
    	File jsonFileName = new File(jsonFile);
    	File targetLocation = new File(output);
    	File languages = new File(languagesfilePath);
    	
    	if(jsonFileName.exists() && !jsonFileName.isDirectory()
    			&& targetLocation.exists() && targetLocation.isDirectory()
    			&& languages.exists() && !languages.isDirectory())
    	{ 
    		
    		//read file path
    		ExcelCodegen codegen = new ExcelCodegen(jsonFile, targetLocation);
    		String excelPath = null;
    		
    		try{
    			LOGGER.info("Processing File --> "+ jsonFile + " For languages specified in " + languages);
    			excelPath = codegen.generateExcel(languagesfilePath);    		
    			LOGGER.info("Sucessfully Created Excel !!!");
    		}
    		catch(Exception e){
    			LOGGER.log(Level.SEVERE, e.getMessage(),e);
    			throw new RuntimeException(e);
    		}
    	}else{
    		throw new RuntimeException("Json File, languages list or Ouput Locations are not valid");
    	}
    	
    	LOGGER.exiting(LOGGER.getName(), "transformJson2Excel");
    }
}