package org.drathveloper.excelToolsCLI.operations;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.drathveloper.excelToolsCLI.CLIUtils.ArgsParser;
import org.drathveloper.excelToolsCLI.model.ParsedSheet;
import java.io.IOException;

public class ParseOperation extends Operation {

    public ParseOperation(ArgsParser parser) throws IOException, OpenXML4JException {
        super(parser);
    }

    @Override
    public void executeOperation(ParsedSheet sheetData) {
        super.closeExcelFile();
        super.writeDataInCSV(sheetData);
    }
}
