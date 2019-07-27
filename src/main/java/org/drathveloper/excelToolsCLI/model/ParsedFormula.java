package org.drathveloper.excelToolsCLI.model;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ParsedFormula {

    private static final String STRICTLY_CELL_REF_PATTERN = "\\$?([A-Z]+)\\$?([0-9]+)";

    private boolean isShared;

    private String formula;

    private int index;

    private String startRange;

    private String endRange;

    public ParsedFormula(){

    }

    public ParsedFormula(ParsedFormula formula){
        this.isShared = formula.isShared();
        this.formula = formula.getFormula();
        this.index = formula.getIndex();
        this.startRange = formula.getStartRange();
        this.endRange = formula.getEndRange();
    }

    public boolean isShared() {
        return isShared;
    }

    public void setShared(boolean shared) {
        isShared = shared;
    }

    public String getFormula(){
        return formula;
    }

    public String getFormula(int rowIndex) {
        return String.format(formula, rowIndex);
    }

    public void setFormula(String formula) {
        Pattern p = Pattern.compile(STRICTLY_CELL_REF_PATTERN);
        Matcher m = p.matcher(formula);
        StringBuffer sb = new StringBuffer(formula.length());
        while(m.find()){
            String match = m.group();
            String tokenized = this.replaceWithPlaceholder(match);
            m.appendReplacement(sb, Matcher.quoteReplacement(tokenized));
        }
        m.appendTail(sb);
        this.formula = sb.toString();
    }

    private String replaceWithPlaceholder(String match){
        StringBuilder str = new StringBuilder();
        for(int i=0; i<match.length(); i++){
            try {
                Integer.parseInt(Character.toString(match.charAt(i)));
                str.append("%1$s");
            } catch(NumberFormatException ex){
                str.append(match.charAt(i));
            }
        }
        return str.toString();
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public String getStartRange() {
        return startRange;
    }

    public void setStartRange(String startRange) {
        this.startRange = startRange;
    }

    public String getEndRange() {
        return endRange;
    }

    public void setEndRange(String endRange) {
        this.endRange = endRange;
    }
}
