package org.drathveloper.excelToolsCLI.model;

import org.drathveloper.excelToolsCLI.handler.AdvancedExcelHandler;
import org.drathveloper.excelToolsCLI.handler.ExcelConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.Month;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.List;

public class ParsedSheet {

    private List<ParsedRow> rawSheetData;

    private String dateFormat;

    private Logger logger = LoggerFactory.getLogger(ParsedSheet.class);

    public ParsedSheet(AdvancedExcelHandler handler){
        rawSheetData = new ArrayList<>(handler.getRawData());
        this.dateFormat = null;
    }

    public void appendData(List<List<String>> rawData){
        logger.info("Started append operation");
        for (List<String> rawDatum : rawData) {
            ParsedRow generatedRow = this.generateRow(rawDatum);
            rawSheetData.add(generatedRow);
        }
        logger.info("Finished append operation");
    }

    public void mergeData(List<List<String>> rawData) {
        logger.info("Started merge operation");
        for (List<String> row : rawData) {
            boolean isSame = false;
            for (ParsedRow prow : rawSheetData) {
                if (prow.getRow().size() == row.size()) {
                    int coincidences = 0;
                    for (int j = 0; j < prow.getRow().size(); j++) {
                        if (prow.getCell(j).getValue().equals(row.get(j))) {
                            coincidences++;
                        }
                    }
                    if (coincidences == prow.getRow().size()) {
                        isSame = true;
                    }
                }
            }
            if (!isSame) {
                ParsedRow generatedRow = this.generateRow(row);
                rawSheetData.add(generatedRow);
            }
        }
        logger.info("Finished merge operation");
    }

    public ParsedRow generateRow(List<String> rawCells){
        ParsedRow row = new ParsedRow();
        ParsedRow previousRow = rawSheetData.get(rawSheetData.size() - 1);
        for(int i=0; i<rawCells.size(); i++){
            ParsedCell cell = new ParsedCell();
            ParsedCell previousCell = previousRow.getCell(i);
            String readValue = rawCells.get(i);
            if(this.isDate(readValue)){
                this.upgradeToDateCell(cell, previousCell);
                readValue = Double.toString(this.convertStringToExcelDate(readValue, dateFormat));
            }
            cell.setValue(readValue);
            row.addCell(cell);
        }
        return row;
    }

    private void upgradeToDateCell(ParsedCell cell, ParsedCell previousCell){
        try {
            if (previousCell != null && !previousCell.isHeaderCell()) { //ToDo: Posibilidad de hoja sin headers
                cell.setStyleRef(previousCell.getStyleRef());
                cell.setTypeRef(previousCell.getTypeRef());
            } else {
                throw new NullPointerException("You shouldnt copy header styles (may differ from real style) nor try to check empty cells");
            }
        } catch(NullPointerException ex){
            int dateStyleIndex = -3;
            cell.setStyleRef(dateStyleIndex);
            cell.setTypeRef(ExcelConstants.NUMBER_TYPE);
        }
    }

    private boolean isDate(String value){
        try {
            SimpleDateFormat df = new SimpleDateFormat(dateFormat);
            df.parse(value);
            return true;
        } catch(ParseException | NullPointerException ex){
            return false;
        }
    }

    private double convertStringToExcelDate(String value, String pattern){
        try {
            SimpleDateFormat df = new SimpleDateFormat(pattern);
            LocalDateTime startDate = LocalDateTime.of(1900, Month.JANUARY, 1, 0, 0, 0);
            LocalTime startTime = LocalTime.of(0,0,0);
            LocalDateTime receivedDate = df.parse(value).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
            double hourDiff = ChronoUnit.SECONDS.between(startTime, receivedDate.toLocalTime());
            //+2 because date is not inclusive and because startDate in Excel is day=0, not 1
            long daysBetween = ChronoUnit.DAYS.between(startDate, receivedDate) + 2;
            double hoursBetween = (hourDiff / 3600) / 24;
            return daysBetween + hoursBetween;
        } catch(ParseException ex){
            return -1;
        }
    }

    public int getRowCount(){
        return rawSheetData.size();
    }

    public List<String> getRowValues(int rowIndex){
        List<String> rowValues = new ArrayList<>();
        for(ParsedCell cell : rawSheetData.get(rowIndex).getRow()){
            rowValues.add(cell.getValue());
        }
        return rowValues;
    }

    public List<ParsedCell> getRow(int rowIndex){
        return rawSheetData.get(rowIndex).getRow();
    }

    public void setDateFormat(String dateFormat){
        this.dateFormat = dateFormat;
    }

}
