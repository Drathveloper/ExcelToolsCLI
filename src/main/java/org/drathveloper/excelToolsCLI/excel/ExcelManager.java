package org.drathveloper.excelToolsCLI.excel;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.xmlbeans.XmlOptions;
import org.drathveloper.excelToolsCLI.handler.AdvancedExcelHandler;
import org.drathveloper.excelToolsCLI.handler.ExcelConstants;
import org.drathveloper.excelToolsCLI.model.ParsedCell;
import org.drathveloper.excelToolsCLI.model.ParsedFormula;
import org.drathveloper.excelToolsCLI.model.ParsedSheet;
import org.openxmlformats.schemas.officeDocument.x2006.relationships.STRelationshipId;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.namespace.QName;
import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

public class ExcelManager {

    private OPCPackage packager;

    private ExcelResources resources;

    private Map<String, String> sheetNames;

    private Logger logger = LoggerFactory.getLogger(ExcelManager.class);

    public ExcelManager(String excelFile) throws OpenXML4JException, IOException {
        this.openFile(excelFile);
        XSSFReader reader = new XSSFReader(packager);
        resources = new ExcelResources(reader);
        this.loadSheetNames();
    }

    private void loadSheetNames() throws OpenXML4JException, IOException {
        sheetNames = new HashMap<>();
        XSSFReader reader = new XSSFReader(packager);
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) reader.getSheetsData();
        while(iter.hasNext()){
            iter.next();
            sheetNames.put(iter.getSheetName(), iter.getSheetPart().getPartName().toString());
        }
    }

    public void openFile(String excelFile) throws InvalidFormatException {
        packager = OPCPackage.open(excelFile);
    }

    public void closeFile() throws IOException{
        packager.close();
    }

    public ParsedSheet readSheet(String sheetName, boolean ignoreFormulas) throws OpenXML4JException, IOException, ParserConfigurationException, SAXException {
        logger.info("Started reading sheet operation");
        ParsedSheet parsedSheet;
        if(sheetNames.get(sheetName) != null){
            XMLReader sheetParser = SAXHelper.newXMLReader();
            PackagePart partSheet = packager.getPartsByName(Pattern.compile(sheetNames.get(sheetName))).get(0);
            InputSource sheetSource = new InputSource(partSheet.getInputStream());
            ContentHandler handler = this.generateContentHandler(ignoreFormulas);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
            parsedSheet = new ParsedSheet((AdvancedExcelHandler) handler);
        } else {
            throw new IllegalArgumentException("Parameter -s must be an existing sheet in the xlsx document");
        }
        logger.info("Finished reading sheet operation");
        return parsedSheet;
    }

    public void writeSheet(ParsedSheet sheet, String sheetName) throws IOException {
        logger.info("Started write sheet operation");
        WorksheetDocument worksheetDocument = WorksheetDocument.Factory.newInstance();
        CTWorksheet worksheet = worksheetDocument.addNewWorksheet();
        CTSheetData sheetData = worksheet.addNewSheetData();
        this.populateSheetData(sheet, sheetData);
        boolean writeResult = resources.writeAllResources();
        if(writeResult){
            this.writeToPart(sheetName, worksheet);
        }
        logger.info("Finished write sheet operation");
    }

    private void populateSheetData(ParsedSheet sheet, CTSheetData sheetData){
        for (int rowNum = 0; rowNum < sheet.getRowCount(); rowNum++) {
            List<ParsedCell> parsedRow = sheet.getRow(rowNum);
            CTRow row = sheetData.addNewRow();
            row.setR(rowNum + 1);
            for (int colNum = 0; colNum < parsedRow.size(); colNum++) {
                if(!parsedRow.get(colNum).isEmptyCell()) {
                    CTCell cell = row.addNewC();
                    this.setCellAttributes(sheet, cell, rowNum, colNum);
                }
            }
        }
    }

    private void setCellAttributes(ParsedSheet sheet, CTCell cell, int rowNum, int colNum){
        ParsedCell parsedCell = sheet.getRow(rowNum).get(colNum);
        String cellRef = CellReference.convertNumToColString(colNum) + (rowNum + 1);
        cell.setR(cellRef);
        this.setCellStyle(cell, parsedCell.getStyleRef());
        this.setCellType(cell, parsedCell.getTypeRef());
        this.setCellValue(cell, parsedCell.getValue());
        this.setCellFormula(cell, parsedCell.getAssociatedFormula(), rowNum);
    }

    private void setCellFormula(CTCell cell, ParsedFormula associatedFormula, int rowNum){
        if (associatedFormula!=null) {
            this.handleCellFormula(cell, associatedFormula, rowNum);
        }
    }

    private void setCellValue(CTCell cell, String value){
        if (cell.getT() == STCellType.S) {
            int sRef = resources.addSharedStringEntry(value);
            cell.setV(Integer.toString(sRef));
        } else {
            if (value!=null) {
                cell.setV(value);
            }
        }
    }

    private void setCellStyle(CTCell cell, int cellStyleIndex){
        if(cellStyleIndex == ExcelConstants.NO_DATE_STYLE){
            int styleIndex = resources.addDateFormatStyle();
            cell.setS(styleIndex);
        } else if(cellStyleIndex >= 0) {
            cell.setS(cellStyleIndex);
        } else {
            //No style
        }
    }

    private void setCellType(CTCell cell, String cellType){
        if(cellType!=null) {
            switch (cellType) {
                case ExcelConstants.BOOLEAN_TYPE:
                    cell.setT(STCellType.B);
                    break;
                case ExcelConstants.ERROR_TYPE:
                    cell.setT(STCellType.E);
                    break;
                case ExcelConstants.FORMULA_TYPE:
                    cell.setT(STCellType.STR);
                    break;
                case ExcelConstants.NUMBER_TYPE:
                    //Do nothing
                    break;
                default:
                    cell.setT(STCellType.S);
                    break;
            }
        } else {
            cell.setT(STCellType.S);
        }
    }

    private void handleCellFormula(CTCell cell, ParsedFormula parsedFormula, int rowNum){
        CTCellFormula formula = cell.addNewF();
        if (parsedFormula.isShared()) {
            formula.setT(STCellFormulaType.SHARED);
            formula.setSi(parsedFormula.getIndex());
            if (parsedFormula.getStartRange() != null && parsedFormula.getEndRange() != null) {
                formula.setRef(parsedFormula.getStartRange() + ExcelConstants.RANGE_SEPARATOR + parsedFormula.getEndRange());
                formula.setStringValue(parsedFormula.getFormula(rowNum + 1));
            }
        } else {
            formula.setStringValue(parsedFormula.getFormula(rowNum + 1));
        }
    }

    private XmlOptions buildXMLOptions(){
        XmlOptions options = new XmlOptions();
        options.setSaveOuter();
        options.setUseDefaultNamespace();
        options.setSaveAggressiveNamespaces();
        options.setCharacterEncoding("UTF-8");
        options.setSaveSyntheticDocumentElement(new QName(CTWorksheet.type.getName().getNamespaceURI(), "worksheet"));
        Map<String, String> relationships = new HashMap<>();
        relationships.put(STRelationshipId.type.getName().getNamespaceURI(), "r");
        options.setSaveSuggestedPrefixes(relationships);
        return options;
    }

    private void writeToPart(String sheetName, CTWorksheet sheet) throws IOException{
        PackagePart sheetPart = packager.getPartsByName(Pattern.compile(sheetNames.get(sheetName))).get(0);
        XmlOptions options = this.buildXMLOptions();
        OutputStream out = sheetPart.getOutputStream();
        sheet.save(out, options);
        out.close();
    }

    private ContentHandler generateContentHandler(boolean ignoreFormula) throws OpenXML4JException, IOException{
        XSSFReader reader = new XSSFReader(packager);
        SharedStringsTable sharedStringsTable = reader.getSharedStringsTable();
        return new AdvancedExcelHandler(sharedStringsTable, ignoreFormula);
    }

    public String formatCellFromStylesTable(ParsedCell cell){
        return resources.formatRawCellContent(cell);
    }

}
