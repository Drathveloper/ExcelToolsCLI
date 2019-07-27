package org.drathveloper.excelToolsCLI.csv;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.CSVWriter;
import org.drathveloper.excelToolsCLI.excel.ExcelManager;
import org.drathveloper.excelToolsCLI.handler.ExcelConstants;
import org.drathveloper.excelToolsCLI.model.ParsedCell;
import org.drathveloper.excelToolsCLI.model.ParsedSheet;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;

public class CSVManager {

    private String csvPath;

    public CSVManager(String csvPath){
        this.csvPath = csvPath;
    }

    public List<List<String>> readCSV(String encoding) throws IOException {
        Reader reader = Files.newBufferedReader(Paths.get(csvPath), Charset.forName(encoding));
        CSVReader csvReader = new CSVReaderBuilder(reader).withSkipLines(1).build();
        String[] nextRecord;
        List<List<String>> rawCSVData = new ArrayList<>();
        while((nextRecord = csvReader.readNext())!= null){
            List<String> records = Arrays.asList(nextRecord);
            rawCSVData.add(records);
        }
        return rawCSVData;
    }

    public void writeCSV(ParsedSheet sheet, ExcelManager excelManager, boolean ignoreFormula) throws IOException {
        FileOutputStream out = new FileOutputStream(csvPath, false);
        this.setBOM(out);
        CSVWriter csvWriter = new CSVWriter(new OutputStreamWriter(out),
                CSVWriter.DEFAULT_SEPARATOR,
                CSVWriter.DEFAULT_QUOTE_CHARACTER,
                CSVWriter.DEFAULT_ESCAPE_CHARACTER,
                CSVWriter.DEFAULT_LINE_END);
        for(int i=0; i < sheet.getRowCount(); i++){
            for(ParsedCell cell : sheet.getRow(i)){
                if(cell.getStyleRef() > 0 && cell.getTypeRef().equals(ExcelConstants.NUMBER_TYPE) && cell.getValue()!=null){
                    cell.setValue(excelManager.formatCellFromStylesTable(cell));
                }
                if(cell.getAssociatedFormula()!=null && !ignoreFormula){
                    cell.setValue(cell.getAssociatedFormula().getFormula());
                }
            }
            List<String> rowValues = sheet.getRowValues(i);
            String[] records = rowValues.toArray(new String[rowValues.size()]);
            csvWriter.writeNext(records);
        }
        csvWriter.close();
    }

    private void setBOM(FileOutputStream out) throws IOException {
        out.write(0xef);
        out.write(0xbb);
        out.write(0xbf);
    }
}
