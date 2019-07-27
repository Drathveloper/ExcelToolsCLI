package org.drathveloper.excelToolsCLI.handler;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.drathveloper.excelToolsCLI.model.ParsedCell;
import org.drathveloper.excelToolsCLI.model.ParsedFormula;
import org.drathveloper.excelToolsCLI.model.ParsedRow;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class AdvancedExcelHandler extends DefaultHandler {

    private static final Pattern CELL_REF_PATTERN = Pattern.compile("(\\$?[A-Z]+)?(\\$?[0-9]+)?", 2);

    private SharedStringsTable sharedStringsTable;

    private boolean isIgnoreFormula;

    private Map<Integer, List<ParsedFormula>> sharedFormulasTable;

    private StringBuffer readValue = new StringBuffer();

    private int rowNum = 0;

    private int lastCol = -1;

    private ParsedFormula formula;

    private List<ParsedRow> rawData;

    private ParsedRow parsedRow;

    private ParsedCell cell;

    public AdvancedExcelHandler(SharedStringsTable sharedStringsTable, boolean isIgnoreFormula){
        this.sharedStringsTable = sharedStringsTable;
        this.isIgnoreFormula = isIgnoreFormula;
        this.rawData = new ArrayList<>();
        this.sharedFormulasTable = new HashMap<>();
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) {
        this.resetReadValue();
        switch(name){
            case ExcelConstants.ROW_TAG:
                String readRowNum = attributes.getValue(ExcelConstants.ROW_NUM_ATTRIBUTE);
                this.initializeRow(readRowNum);
                break;
            case ExcelConstants.CELL_TAG:
                String cellRef = attributes.getValue(ExcelConstants.CELL_REF_ATTRIBUTE);
                String styleRef = attributes.getValue(ExcelConstants.CELL_STYLE_ATTRIBUTE);
                String cellType = attributes.getValue(ExcelConstants.CELL_TYPE_ATTRIBUTE);
                this.initializeCell(cellRef, styleRef, cellType);
                break;
            case ExcelConstants.FORMULA_TAG:
                String formulaType = attributes.getValue(ExcelConstants.FORMULA_TYPE_ATTRIBUTE);
                if(formulaType != null && formulaType.equals(ExcelConstants.SHARED_TYPE)){
                    String refValue = attributes.getValue(ExcelConstants.FORMULA_REF_ATTRIBUTE);
                    String indexValue = attributes.getValue(ExcelConstants.FORMULA_INDEX_ATTRIBUTE);
                    if(refValue!=null){
                        formula = new ParsedFormula();
                        formula.setShared(true);
                        formula.setIndex(Integer.parseInt(indexValue));
                        formula.setStartRange(refValue.split(":")[0]);
                        formula.setEndRange(refValue.split(":")[1]);
                    } else {
                        formula = new ParsedFormula();
                        formula.setShared(true);
                        formula.setIndex(Integer.parseInt(indexValue));
                    }
                } else {
                    formula = new ParsedFormula();
                    formula.setShared(false);
                    formula.setIndex(-1);
                    formula.setStartRange(null);
                    formula.setEndRange(null);
                }
                break;
        }
    }

    private int getColFromReference(String reference) {
        int plingPos = reference.lastIndexOf(33);
        String cell = reference.substring(plingPos + 1).toUpperCase(Locale.ROOT);
        Matcher matcher = CELL_REF_PATTERN.matcher(cell);
        if (!matcher.matches()) {
            throw new IllegalArgumentException("Invalid CellReference: " + reference);
        } else {
            String col = matcher.group(1);
            return CellReference.convertColStringToIndex(col);
        }
    }

    private void initializeRow(String readRowNum) {
        rowNum = Integer.parseInt(readRowNum);
        parsedRow = new ParsedRow();
    }

    private void initializeCell(String cellRef, String styleRef, String cellType) {
        int colNum = this.getColFromReference(cellRef);
        this.addMissingCols(colNum);
        lastCol = colNum;
        cell = new ParsedCell();
        try {
            int styleIndex = Integer.parseInt(styleRef);
            cell.setStyleRef(styleIndex);
        } catch(NumberFormatException ex){
            cell.setStyleRef(ExcelConstants.NO_STYLE);
        }
        if(cellType==null) {
            cell.setTypeRef(ExcelConstants.NUMBER_TYPE);
        } else {
            cell.setTypeRef(cellType);
        }
        cell.setHeaderCell(rowNum <= 1);
    }

    private void resetReadValue(){
        readValue.setLength(0);
    }

    private void addMissingCols(int colNum){
        int missingCols = calculateMissingCols(colNum);
        for(int i=0; i<missingCols; i++){
            parsedRow.addCell(new ParsedCell());
        }
    }

    private int calculateMissingCols(int colNum){
        return colNum - lastCol - 1;
    }

    @Override
    public void endElement(String uri, String localName, String name) {
        switch(name){
            case ExcelConstants.CELL_TAG:
                this.finalizeCell();
                break;
            case ExcelConstants.CELL_VALUE_TAG:
                this.handleCellType(readValue.toString());
                break;
            case ExcelConstants.FORMULA_TAG:
                if(formula.isShared()){
                    if(formula.getStartRange() != null && formula.getEndRange()!=null){
                        formula.setFormula(readValue.toString());
                        ParsedFormula sharedFormula = new ParsedFormula(formula);
                        if(sharedFormulasTable.get(lastCol)==null){
                            List<ParsedFormula> formulaList = new ArrayList<>();
                            formulaList.add(sharedFormula);
                            sharedFormulasTable.put(lastCol, formulaList);
                        } else {
                            sharedFormulasTable.get(lastCol).add(sharedFormula);
                        }
                    } else {
                        if(sharedFormulasTable.get(lastCol)!=null){
                            for(ParsedFormula formulaFromList : sharedFormulasTable.get(lastCol)){
                                int col = this.getColFromReference(formulaFromList.getEndRange());
                                if(col == lastCol){
                                    if(formulaFromList.getIndex() == formula.getIndex()){
                                        formula.setFormula(formulaFromList.getFormula());
                                        break;
                                    }
                                }
                            }
                        }
                    }
                } else {
                    formula.setFormula(readValue.toString());
                }
                cell.setAssociatedFormula(formula);
                break;
            case ExcelConstants.ROW_TAG:
                this.finalizeRow();
                break;
        }
        this.resetReadValue();
    }

    private void handleCellType(String reference){
        if(ExcelConstants.SST_STRING_TYPE.equals(cell.getTypeRef())) {
            this.addCellValueFromSharedStrings(reference);
        } else if(ExcelConstants.NUMBER_TYPE.equals(cell.getTypeRef())){
            cell.setValue(readValue.toString());
        }
    }

    private void finalizeCell(){
        parsedRow.addCell(new ParsedCell(cell));
        cell = new ParsedCell();
    }

    private void addCellValueFromSharedStrings(String reference){
        RichTextString str = sharedStringsTable.getItemAt(Integer.parseInt(reference));
        cell.setValue(str.getString());
    }

    private void finalizeRow(){
        if(!parsedRow.isEmptyRow(isIgnoreFormula)) {
            rawData.add(new ParsedRow(parsedRow));
        }
        parsedRow = new ParsedRow();
        lastCol = -1;
    }

    @Override
    public void characters(char[] ch, int start, int length) {
        readValue.append(ch, start, length);
    }

    public List<ParsedRow> getRawData(){
        return rawData;
    }
}
