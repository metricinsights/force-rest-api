package com.force.api.util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;

public class XLSXToCSVConverter {
    private static final Logger logger = LoggerFactory.getLogger(XLSXToCSVConverter.class);

    private Workbook workbook = null;
    private DataFormatter formatter = null;
    private FormulaEvaluator evaluator = null;

    private ArrayList<ArrayList<String>> csvData = null;
    private int maxRowWidth = 0;

    private static final int FORMATTING_CONVENTION = 0;
    private static final int EXCEL_STYLE_ESCAPING = 0;
    private static final String DEFAULT_SEPARATOR = ",";
    private static final String SEPARATOR = DEFAULT_SEPARATOR;


    public InputStream convertToCSV(InputStream stream) throws IOException {
        logger.info("Converting XLSX to CSV...");
        openWorkbook(stream);
        convertToCSV();
        ByteArrayOutputStream outputStream = getCSVStream();
        return new ByteArrayInputStream(outputStream.toByteArray());
    }

    private void openWorkbook(InputStream stream) throws IOException {
        logger.info("Open XLSX workbook");
        try {
            this.workbook = new XSSFWorkbook(stream);
            this.evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
            this.formatter = new DataFormatter(true);
        } finally {
            if (stream != null) {
                stream.close();
            }
        }
    }

    private void convertToCSV() {
        Sheet sheet;
        Row row;
        int lastRowNum;
        this.csvData = new ArrayList<>();

        // Discover how many sheets there are in the workbook....
        int numSheets = this.workbook.getNumberOfSheets();
        logger.info("There are {} sheets in workbook", numSheets);

        for (int i = 0; i < numSheets; i++) {
            sheet = this.workbook.getSheetAt(i);
            if (sheet.getPhysicalNumberOfRows() > 0) {
                // Note down the index number of the bottom-most row and
                // then iterate through all of the rows on the sheet starting
                // from the very first row - number 1 - even if it is missing.
                // Recover a reference to the row and then call another method
                // which will strip the data from the cells and build lines
                // for inclusion in the resylting CSV file.
                lastRowNum = sheet.getLastRowNum();
                logger.info("{}-Sheet: {} rows", i, lastRowNum+1);
                for (int j = 0; j <= lastRowNum; j++) {
                    row = sheet.getRow(j);
                    rowToCSV(row);
                }
            }
        }
    }

    private void rowToCSV(Row row) {
        Cell cell;
        int lastCellNum;
        ArrayList<String> csvLine = new ArrayList<>();

        // Check to ensure that a row was recovered from the sheet as it is
        // possible that one or more rows between other populated rows could be
        // missing - blank. If the row does contain cells then...
        if (row != null) {

            // Get the index for the right most cell on the row and then
            // step along the row from left to right recovering the contents
            // of each cell, converting that into a formatted String and
            // then storing the String into the csvLine ArrayList.
            lastCellNum = row.getLastCellNum();
            for (int i = 0; i <= lastCellNum; i++) {
                cell = row.getCell(i);
                if (cell == null) {
                    csvLine.add("");
                } else {
                    if (cell.getCellType() != CellType.FORMULA) {
                        csvLine.add(this.formatter.formatCellValue(cell));
                    } else {
                        csvLine.add(this.formatter.formatCellValue(cell, this.evaluator));
                    }
                }
            }
            // Make a note of the index number of the right most cell. This value
            // will later be used to ensure that the matrix of data in the CSV file
            // is square.
            if (lastCellNum > this.maxRowWidth) {
                this.maxRowWidth = lastCellNum;
            }
        }
        this.csvData.add(csvLine);
    }

    private ByteArrayOutputStream getCSVStream() throws IOException {
        ArrayList<String> line ;
        StringBuilder builder;
        String csvLineElement;
        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
             BufferedWriter bw = new BufferedWriter(new OutputStreamWriter(outputStream))){

            for (int i = 0; i < this.csvData.size(); i++) {
                builder = new StringBuilder();

                line = this.csvData.get(i);
                for (int j = 0; j < this.maxRowWidth; j++) {
                    if (line.size() > j) {
                        csvLineElement = line.get(j);
                        if (csvLineElement != null) {
                            builder.append(this.escapeEmbeddedCharacters(csvLineElement));
                        }
                    }
                    if (j < (this.maxRowWidth - 1)) {
                        builder.append(SEPARATOR);
                    }
                }

                bw.write(builder.toString().trim());

                // Condition the inclusion of new line characters so as to
                // avoid an additional, superfluous, new line at the end of
                // the file.
                if (i < (this.csvData.size() - 1)) {
                    bw.newLine();
                }
            }
            logger.info("Wrote {} rows to CSV stream", this.csvData.size());
            return outputStream;
        }
    }

    private String escapeEmbeddedCharacters(String field) {
        StringBuilder builder;

        if (FORMATTING_CONVENTION == EXCEL_STYLE_ESCAPING) {

            // Firstly, check if there are any speech marks (") in the field;
            // each occurrence must be escaped with another set of spech marks
            // and then the entire field should be enclosed within another
            // set of speech marks. Thus, "Yes" he said would become
            // """Yes"" he said"
            if (field.contains("\"")) {
                builder = new StringBuilder(field.replace("\"", "\\\"\\\""));
                builder.insert(0, "\"");
                builder.append("\"");
            } else {
                // If the field contains either embedded separator or EOL
                // characters, then escape the whole field by surrounding it
                // with speech marks.
                builder = new StringBuilder(field);
                if ((builder.indexOf(SEPARATOR)) > -1 || (builder.indexOf("\n")) > -1) {
                    builder.insert(0, "\"");
                    builder.append("\"");
                }
            }
            return (builder.toString().trim());
        }
        // The only other formatting convention this class obeys is the UNIX one
        // where any occurrence of the field separator or EOL character will
        // be escaped by preceding it with a backslash.
        else {
            if (field.contains(SEPARATOR)) {
                field = field.replace(SEPARATOR, ("\\\\" + SEPARATOR));
            }
            if (field.contains("\n")) {
                field = field.replace("\n", "\\\\\n");
            }
            return (field);
        }
    }
}