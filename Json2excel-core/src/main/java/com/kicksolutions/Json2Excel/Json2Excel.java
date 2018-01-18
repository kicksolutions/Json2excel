package com.kicksolutions.Json2Excel;

import java.util.logging.Logger;

import com.kicksolutions.CliArgs;
import com.kicksolutions.Json2Excel.excel.ExcelGenerator;



/**
 * MBARDIYA
 *
 */
public class Json2Excel
{
	private static final Logger LOGGER = Logger.getLogger(Json2Excel.class.getName());
	private static final String USAGE = new StringBuilder()
			.append(" Usage: ")
			.append(Json2Excel.class.getName()).append(" <options> \n")
			.append(" -i <json file> ")
			.append(" -o <output directory> ")
			.append(" -l <Language file>").toString();
	
	public Json2Excel() {
		super();
	}
	
	/**
	 * 
	 * @param args
	 */
    public static void main( String[] args )
    {
    	Json2Excel json2excel = new Json2Excel();
    	json2excel.init(args);   	
    }
    

	/**
     * 
     * @param args
     */
    private void init(String args[]){
    	LOGGER.entering(LOGGER.getName(), "init");
    	
    	CliArgs cliArgs = new CliArgs(args);
    	String jsonFile = cliArgs.getArgumentValue("-i", "");
    	String output = cliArgs.getArgumentValue("-o","");
    	String languagesFilePath = cliArgs.getArgumentValue("-l","");
    	
    	if(!(jsonFile.isEmpty() && output.isEmpty()))
    	{
      		process(jsonFile, output, languagesFilePath);
    	}
    	else{
    		LOGGER.severe(USAGE);
    	}
    	
    	LOGGER.exiting(LOGGER.getName(), "init");
    }
    
    /**
     * 
     * @param specFile
     * @param output
     */
    private void process(String jsonFile,String output, String languagesfilePath){
    	ExcelGenerator generator = new ExcelGenerator();
    	generator.transformJson2Excel(jsonFile, output,languagesfilePath);
    }    
}