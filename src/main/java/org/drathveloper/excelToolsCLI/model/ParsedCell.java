package org.drathveloper.excelToolsCLI.model;

import org.drathveloper.excelToolsCLI.handler.ExcelConstants;

public class ParsedCell {

    private int styleRef;

    private String typeRef;

    private String value;

    private ParsedFormula associatedFormula;

    private boolean headerCell;

    private String formatString;

    public ParsedCell(int styleRef, String typeRef, String value, ParsedFormula associatedFormula, int rowNum, String formatString) {
        this.styleRef = styleRef;
        this.typeRef = typeRef;
        this.value = value;
        this.associatedFormula = associatedFormula;
        this.headerCell = rowNum > 0;
        this.formatString = formatString;
    }

    public ParsedCell(ParsedCell cell){
        this.styleRef = cell.getStyleRef();
        this.typeRef = cell.getTypeRef();
        this.value = cell.getValue();
        this.associatedFormula = cell.associatedFormula;
        this.headerCell = cell.isHeaderCell();
        this.formatString = cell.getFormatString();
    }

    public ParsedCell(){
        this.styleRef = -2;
        this.typeRef = null;
        this.value = null;
        this.associatedFormula = null;
        this.headerCell = false;
        this.formatString = null;
    }

    public int getStyleRef() {
        return styleRef;
    }

    public String getTypeRef() {
        return typeRef;
    }

    public String getValue() {
        return value;
    }

    public void setStyleRef(int styleRef) {
        this.styleRef = styleRef;
    }

    public void setTypeRef(String typeRef) {
        this.typeRef = typeRef;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public boolean isFormulaCell(){
        return associatedFormula!=null;
    }

    public ParsedFormula getAssociatedFormula() {
        return associatedFormula;
    }

    public void setAssociatedFormula(ParsedFormula associatedFormula) {
        this.associatedFormula = associatedFormula;
    }

    public boolean isEmptyCell(){
        return (styleRef == -2 && typeRef == null && value == null && associatedFormula == null) ||
                (styleRef >= -1 && typeRef.equals(ExcelConstants.NUMBER_TYPE) && value==null && associatedFormula==null);
    }

    public boolean isHeaderCell() {
        return headerCell;
    }

    public void setHeaderCell(boolean headerCell) {
        this.headerCell = headerCell;
    }

    public String getFormatString() {
        return formatString;
    }

    public void setFormatString(String formatString) {
        this.formatString = formatString;
    }
}
