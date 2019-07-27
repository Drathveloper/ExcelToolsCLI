package org.drathveloper.excelToolsCLI.operations;

import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.drathveloper.excelToolsCLI.CLIUtils.ArgsParser;
import org.drathveloper.excelToolsCLI.model.ParsedSheet;
import java.io.IOException;
import java.util.List;

public class AppendOperation extends Operation {

    public AppendOperation(ArgsParser parser) throws IOException, OpenXML4JException {
        super(parser);
    }

    @Override
    public void executeOperation(ParsedSheet sheetData) {
        List<List<String>> readCSV = super.readCSVData();
        sheetData.appendData(readCSV);
        super.writeDataInExcel(sheetData);
    }

}
