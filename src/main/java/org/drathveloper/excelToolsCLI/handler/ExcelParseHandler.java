package org.drathveloper.excelToolsCLI.handler;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelParseHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    private Map<Integer, List<String>> rawSheetData;

    private List<String> rowData;

    private int lastCol;

    private int lastRow;

    public ExcelParseHandler(){
        rawSheetData = new HashMap<>();
        rowData = new ArrayList<>();
        lastCol = -1;
        lastRow = -1;
    }

    @Override
    public void startRow(int rowNum) {
        lastCol = -1;
        lastRow = rowNum;
    }

    @Override
    public void endRow(int rowNum) {
        if(!this.checkEmptyRow()){
            rawSheetData.put(rowNum, new ArrayList<>(rowData));
        }
        rowData = new ArrayList<>();
    }

    @Override
    public void cell(String cellRef, String cellValue, XSSFComment cellComment) {
        if(cellRef == null) {
            cellRef = new CellAddress(lastRow, lastCol).formatAsString();
        }
        int currentCol = (new CellReference(cellRef)).getCol();
        this.addMissingCols(currentCol);
        rowData.add(cellValue);
    }

    private boolean checkEmptyRow(){
        for(String s : rowData){
            if( s!=null && !s.equals("") ){
                return false;
            }
        }
        return true;
    }

    private void addMissingCols(int currentCol){
        int missedCols = currentCol - lastCol - 1;
        for(int i=0; i<missedCols; i++){
            rowData.add("");
        }
        lastCol = currentCol;
    }

    public Map<Integer, List<String>> getRawData(){
        return rawSheetData;
    }
}
