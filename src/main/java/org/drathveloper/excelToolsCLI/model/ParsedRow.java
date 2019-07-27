package org.drathveloper.excelToolsCLI.model;

import java.util.ArrayList;
import java.util.List;

public class ParsedRow {

    private List<ParsedCell> cellList;

    public ParsedRow(ParsedRow row){
        cellList = new ArrayList<>(row.getRow());
    }

    public ParsedRow(){
        this.cellList = new ArrayList<>();
    }

    public void addCell(ParsedCell cell){
        cellList.add(cell);
    }

    public ParsedCell getCell(int colIndex){
        return cellList.get(colIndex);
    }

    public List<ParsedCell> getRow(){
        return cellList;
    }

    public void setRow(List<ParsedCell> cellList){
        this.cellList = cellList;
    }

    public boolean isEmptyRow(boolean ignoreFormula){
        for(ParsedCell cell : cellList){
            if(!cell.isEmptyCell()){
                if(cell.isFormulaCell()){
                    return ignoreFormula;
                }
                return false;
            }
        }
        return true;
    }

    public int getColSize(){
        return cellList.size();
    }
}
