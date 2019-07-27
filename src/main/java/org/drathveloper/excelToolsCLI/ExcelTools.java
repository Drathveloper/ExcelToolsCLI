package org.drathveloper.excelToolsCLI;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.drathveloper.excelToolsCLI.CLIUtils.ArgsParser;
import org.drathveloper.excelToolsCLI.CLIUtils.ParameterList;
import org.drathveloper.excelToolsCLI.CLIUtils.ValidOperations;
import org.drathveloper.excelToolsCLI.operations.AppendOperation;
import org.drathveloper.excelToolsCLI.operations.MergeOperation;
import org.drathveloper.excelToolsCLI.operations.Operation;
import org.drathveloper.excelToolsCLI.operations.ParseOperation;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import java.io.IOException;

public class ExcelTools implements Runnable {

    private String[] args;

    private Logger logger = LoggerFactory.getLogger(ExcelTools.class);

    private static final String[] requiredParameters = {ParameterList.EXCEL_PATH, ParameterList.CSV_PATH, ParameterList.OPERATION_MODE, ParameterList.SHEET_NAME};

    private static final String[] optionalParameters = {ParameterList.DATE_FORMAT, ParameterList.IGNORE_FORMULA, ParameterList.CSV_ENCODING, ParameterList.HELP_COMMAND};

    public ExcelTools(String[] args){
        this.args = args;
    }

    @Override
    public void run() {
        if(args!=null){
            try {
                ArgsParser parser = new ArgsParser(args, requiredParameters, optionalParameters);
                String parsedOperation = parser.getParametersFromOptionAsString(ParameterList.OPERATION_MODE);
                ValidOperations selectedOperation = this.getOperationOption(parsedOperation);
                Operation operation;
                if(selectedOperation.equals(ValidOperations.APPEND)){
                    operation = new AppendOperation(parser);
                } else if(selectedOperation.equals(ValidOperations.MERGE)){
                    operation = new MergeOperation(parser);
                } else if(selectedOperation.equals(ValidOperations.PARSE)){
                    operation = new ParseOperation(parser);
                } else {
                    throw new IllegalArgumentException("Operation mode is not supported");
                }
                operation.execute();
            } catch(IllegalArgumentException ex){
                logger.info("Problem parsing arguments. Probably one or many arguments are unrecognized\nMessage: " + ex.getMessage());
            } catch(OpenXML4JException | IOException ex){
                logger.info("There was a problem opening Excel file. You should ensure that the file exists and its a valid xlsx file\n" +
                        "Message: " + ex.getMessage());
            }
        }
    }

    private ValidOperations getOperationOption(String operation) throws IllegalArgumentException {
        switch(operation.toLowerCase()){
            case "merge":
                return ValidOperations.MERGE;
            case "append":
                return ValidOperations.APPEND;
            case "parse":
                return ValidOperations.PARSE;
            default:
                throw new IllegalArgumentException("Bad operation mode option");
        }
    }

}
