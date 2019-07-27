package org.drathveloper.excelToolsCLI.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.drathveloper.excelToolsCLI.model.ParsedCell;

import java.io.IOException;
import java.io.OutputStream;

public class ExcelResources {

    private static final int DEFAULT_DATE_FORMAT = 22;

    private SharedStringsTable sharedStringsTable;

    private StylesTable stylesTable;

    public ExcelResources(XSSFReader reader) throws InvalidFormatException, IOException {
        sharedStringsTable = reader.getSharedStringsTable();
        stylesTable = reader.getStylesTable();
    }

    public int addDateFormatStyle(){
        XSSFCellStyle cellStyle = stylesTable.createCellStyle();
        cellStyle.setDataFormat(DEFAULT_DATE_FORMAT);
        cellStyle.setFillPattern(FillPatternType.NO_FILL);
        return cellStyle.getIndex();
    }

    public int addSharedStringEntry(String value){
        RichTextString entry = new XSSFRichTextString(value);
        return sharedStringsTable.addSharedStringItem(entry);
    }

    public String formatRawCellContent(ParsedCell cell){
        try {
            DataFormatter formatter = new DataFormatter();
            XSSFCellStyle cellStyle = stylesTable.getStyleAt(cell.getStyleRef());
            int dataFormatIndex = cellStyle.getDataFormat();
            String dataFormatString = cellStyle.getDataFormatString();
            return formatter.formatRawCellContents(Double.parseDouble(cell.getValue()), dataFormatIndex, dataFormatString);
        } catch(NumberFormatException | NullPointerException ex){
            return null;
        }
    }

    private boolean writeSharedStringsTable() {
        OutputStream out;
        try {
            out = sharedStringsTable.getPackagePart().getOutputStream();
            sharedStringsTable.writeTo(out);
            out.close();
        } catch(IOException ex){
            return false;
        }
        return true;
    }

    private boolean writeStylesTable() {
        OutputStream out;
        try {
            out = stylesTable.getPackagePart().getOutputStream();
            stylesTable.writeTo(out);
            out.close();
        } catch(IOException ex){
            return false;
        }
        return true;
    }

    public boolean writeAllResources(){
        return this.writeSharedStringsTable() && this.writeStylesTable();
    }
}
