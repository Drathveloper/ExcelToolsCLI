package org.drathveloper.excelToolsCLI.operations;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.drathveloper.excelToolsCLI.CLIUtils.ArgsParser;
import org.drathveloper.excelToolsCLI.CLIUtils.ParameterList;
import org.drathveloper.excelToolsCLI.csv.CSVManager;
import org.drathveloper.excelToolsCLI.excel.ExcelManager;
import org.drathveloper.excelToolsCLI.model.ParsedSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.util.List;

public abstract class Operation {

    private Logger logger = LoggerFactory.getLogger(Operation.class);

    private ExcelManager excelManager;

    private CSVManager csvManager;

    private String sheetName;

    private String dateFormat;

    private String csvEncoding;

    private boolean ignoreFormula;

    public Operation(ArgsParser parser) throws IOException, OpenXML4JException {
        String parsedDateFormat = parser.getParametersFromOptionAsString(ParameterList.DATE_FORMAT);
        String parsedEncoding = parser.getParametersFromOptionAsString(ParameterList.CSV_ENCODING);
        excelManager = new ExcelManager(parser.getParametersFromOptionAsString(ParameterList.EXCEL_PATH));
        csvManager = new CSVManager(parser.getParametersFromOptionAsString(ParameterList.CSV_PATH));
        ignoreFormula = parser.isOptionInArgs(ParameterList.IGNORE_FORMULA);
        sheetName = parser.getParametersFromOptionAsString(ParameterList.SHEET_NAME);
        dateFormat = (parsedDateFormat.equals("") ? ParameterList.DEFAULT_DATE_FORMAT : parsedDateFormat);
        this.csvEncoding = (parsedEncoding.equals("") ? ParameterList.DEFAULT_ENCODING : parsedEncoding);
    }

    public void execute(){
        ParsedSheet sheetData = this.readExcelData();
        this.executeOperation(sheetData);
    }

    protected ParsedSheet readExcelData(){
        try {
            logger.info("Started reading Excel data");
            ParsedSheet sheetData = excelManager.readSheet(sheetName, ignoreFormula);
            sheetData.setDateFormat(dateFormat);
            logger.info("Finished reading Excel data");
            return sheetData;
        } catch (IOException | OpenXML4JException | SAXException | ParserConfigurationException ex){
            logger.info("There was a problem opening Excel file. You should ensure that the file exists and its a valid xlsx file\n" +
                    "Message: " + ex.getMessage());
        }
        return null;
    }

    protected List<List<String>> readCSVData() {
        try {
            logger.info("Started reading CSV data");
            List<List<String>> readCSV = csvManager.readCSV(csvEncoding);
            logger.info("Finished reading CSV data");
            return readCSV;
        } catch(IOException ex){
            logger.info("There was a problem parsing CSV file.\n" +
                    "Message: " + ex.getMessage());
        }
        return null;
    }

    protected void writeDataInExcel(ParsedSheet sheetData) {
        try {
            logger.info("Started writing data in Excel");
            excelManager.writeSheet(sheetData, sheetName);
            this.closeExcelFile();
            logger.info("Finished writing data in Excel");
        } catch(IOException ex){
            logger.info("There was a problem writing Excel file.\n" +
                    "Message: " + ex.getMessage());
        }
    }

    protected void writeDataInCSV(ParsedSheet sheetData) {
        try {
            csvManager.writeCSV(sheetData, excelManager, ignoreFormula);
        } catch (IOException ex){
            logger.info("There was a problem writing CSV file.\n" +
                    "Message: " + ex.getMessage());
        }
    }

    protected void closeExcelFile(){
        try {
            excelManager.closeFile();
        } catch(IOException ex){
            logger.info("There was a problem closing Excel file.\n" +
                    "Message: " + ex.getMessage());
        }
    }

    protected abstract void executeOperation(ParsedSheet sheetData);

}
