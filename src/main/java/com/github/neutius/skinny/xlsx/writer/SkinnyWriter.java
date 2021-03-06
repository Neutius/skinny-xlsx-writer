package com.github.neutius.skinny.xlsx.writer;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

/**
 * Extremely simple and light-weight writer for .xlsx files. The basic use flow is as follows:
 * <ol>
 * <li>The basic constructor immediately writes a .xlsx file with the specified name to the specified target folder.
 * This file will have a single sheet with the specified name and no content.</li>
 * <li>A single column header row can be added to each sheet, and zero or more content rows can be added.</li>
 * <li>New sheets can be added, and to each new sheet column headers and content can be added.</li>
 * <li>At any point, the current in memory data can be written to the .xlsx file, writing over any previously written .xlsx file.</li>
 * <li>Closing this class is unnecessary: the in memory workbook has nothing to close, and any OutputStream used is immediately
 * closed after use.</li>
 * </ol>
 * <p>
 * Besides the basic methods to add sheets, column headers and content, several convenience methods are available that call on one
 * or more of the basic methods.
 * <p>
 * This class only moves forward: column headers can only be added first to a sheet, any content row is added to the bottom of the
 * current sheet, and only the current sheet can be manipulated in any way.
 * Data (sheets, column headers, content) can only be added using this class. There are no methods for removing or deleting data.
 * There are no methods for accessing or updating any added data.
 * <p>
 * This class is basically a thin convenience wrapper around Apache POI's XSSF User Model, allowing other classes to write in memory
 *  data from Java code to a .xlsx file without interacting with Apache POI directly and with a minimum amount of method calls.
 * <p>
 * This class is made final to prevent sub-classing. Copying and adjusting the source code is permitted.
 */

public final class SkinnyWriter {

    private final File targetFile;
    private final CellStyle columnHeaderCellStyle;


    private XSSFWorkbook workbook;
    private XSSFCellStyle currentCellStyle;
    private XSSFFont columnHeaderFont = new XSSFFont();
    private XSSFSheet currentSheet;
    private int currentColumnAmount;
    private int rowIndex;

    {
        columnHeaderFont.setBold(true);
    }

    /**
     * This static method functions as a one-stop shop for all SkinnyWriter functionality. With a single method call:
     * <ol>
     *     <li>A SkinnyWriter instance with an in memory workbook is created.</li>
     *     <li>One sheet is added to the workbook for each entry in the <code>List&lt;SkinnySheetContent&gt;</code> passed in.</li>
     *     <li>A .xlsx file is written to the target folder.</li>
     * </ol>
     * <p>
     * This method is the ultimate convenience method of the SkinnyWriter class, calling (directly or indirectly) several other
     * convenience methods and every basic method of this class.
     *
     * @param targetFolder     The target location for the .xlsx file
     * @param fileName         The base name of the .xlsx that will be written.
     *                         No extension needed, this constructor automatically adds the .xlsx extension, without checking if an
     *                         extension is already present, e.g. passing in "myFile.xlsx" will result in a file named "myFile.xlsx.xlsx".
     *                         If null or an empty String is passed in, the file will be given a name.
     * @param sheetContentList A List of objects implementing the SkinnySheetContent interface.
     *                         Each object in the List represents a sheet to be added to the .xlsx file.
     * @throws IOException Any Exception occurring while writing to the file system will remain uncaught.
     */

    public static void writeContentToFileSystem(File targetFolder, String fileName, List<SkinnySheetContent> sheetContentList)
            throws IOException {
        SkinnyWriter writer = new SkinnyWriter(targetFolder, fileName);
        writer.addSeveralSheetsToWorkbook(sheetContentList);
        writer.writeToFile();
    }

    /**
     * Calling this constructor will initialize an in memory workbook with no sheets. A sheet has to be added before any column
     * headers or content rows can be added, and before a valid .xlsx can be written to the file system.
     * <p>
     * This constructor does not write anything to the file system.
     *
     * @param targetFolder The target directory where the .xlsx file will be written to. Must be an existing directory.
     * @param fileName     The base name of the .xlsx that will be written.
     *                     No extension needed, this constructor automatically adds the .xlsx extension, without checking if an
     *                     extension is already present, e.g. passing in "myFile.xlsx" will result in a file named "myFile.xlsx.xlsx".
     *                     If null or an empty String is passed in, the file will be given a name.
     */

    public SkinnyWriter(File targetFolder, String fileName) {
        targetFile = new File(targetFolder, SkinnyUtil.sanitizeFileName(fileName) + SkinnyUtil.EXTENSION);
        workbook = new XSSFWorkbook();
        columnHeaderCellStyle = SkinnyUtil.createColumnHeaderCellStyle(workbook);
    }

    /**
     * Calling this constructor will initialize an in memory workbook with a single empty Sheet, which will be
     * immediately written to disk. This ensures any IOException will be encountered as quickly as possible.
     * <p>
     * Warning: if the target directory already has a .xlsx file with the same base name,
     * it will be overwritten with no further warning.
     *
     * @param targetFolder   The target directory where the .xlsx file will be written to. Must be an existing directory.
     * @param fileName       The base name of the .xlsx that will be written.
     *                       No extension needed, this constructor automatically adds the .xlsx extension, without checking if an
     *                       extension is already present, e.g. passing in "myFile.xlsx" will result in a file named "myFile.xlsx.xlsx".
     *                       If null or an empty String is passed in, the file will be given a name.
     * @param firstSheetName The name of the first sheet of the .xlsx file.
     *                       If the passed in String consists of more than 31 characters, only the first 31 characters will be used.
     *                       If null or an empty String is passed in, the sheet will be given a name.
     * @throws IOException Any Exception that occurs while creating a file on the file system or writing to this file
     *                     will remain uncaught.
     */

    public SkinnyWriter(File targetFolder, String fileName, String firstSheetName) throws IOException {
        targetFile = new File(targetFolder, SkinnyUtil.sanitizeFileName(fileName) + SkinnyUtil.EXTENSION);
        workbook = new XSSFWorkbook();
        columnHeaderCellStyle = SkinnyUtil.createColumnHeaderCellStyle(this.workbook);
        createNewSheet(firstSheetName);
        writeToFile();
    }


    /**
     * Creates a row at the top of the current sheet, and adds the parameter values as cell values to this row.
     * <p>
     * This row is unique in two ways:
     * <ol>
     * <li>Column header text will be given a bold font.</li>
     * <li>A freeze pane will be applied to the column header row.</li>
     * </ol>
     * <p>
     * Adding column headers to a sheet is optional.
     * <p>
     * Column headers should be added first and only once:
     * <ul>
     * <li>Any attempt to add column headers to a non-empty sheet will result in an IllegalStateException.</li>
     * </ul>
     * <p>
     * Column header text should consist of at least 1 non white space character:
     * <ul>
     * <li>If the parameter contains any null values, a NullPointerException will occur and remain uncaught.</li>
     * <li>If the parameter contains any blank Strings, an IllegalArgumentException will be thrown.</li>
     * </ul>
     * <p>
     * This is a basic method, that is called by several other methods.
     *
     * @param columnHeaderRow The List of String values to be added to the column header row.
     * @throws NullPointerException     Passing in any null value will result in a NullPointerException, which will remain uncaught.
     *                                  Calling this method before any sheet has been added will result in the same Exception.
     * @throws IllegalArgumentException Will be thrown when any <code>String</code> value is blank, i.e., an empty String or
     *                                  a String consisting of nothing but white space characters.
     * @throws IllegalStateException    Will be thrown when the current sheet is not empty.
     */

    public void addColumnHeaderRowToCurrentSheet(List<String> columnHeaderRow) {
        if (currentSheet.getRow(0) != null) {
            throw new IllegalStateException("Column headers should be added first, and should be added only once.");
        }
        if (columnHeaderRow.stream().anyMatch(String::isBlank)) {
            throw new IllegalArgumentException("Column header text should not be blank");
        }

        XSSFRow headerColumnRow = currentSheet.createRow(rowIndex++);

        for (int columnIndex = 0; columnIndex < columnHeaderRow.size(); columnIndex++) {
            XSSFCell currentCell = headerColumnRow.createCell(columnIndex);
            currentCell.setCellValue(columnHeaderRow.get(columnIndex));
            currentCell.setCellStyle(columnHeaderCellStyle);
        }

        currentSheet.createFreezePane(0, 1);
    }

    /**
     * Creates a new row at the bottom of the current sheet, and adds the parameter values (if any) as cell values to this row.
     * <p>
     * This is a basic method, that is called by several other methods.
     *
     * @param rowContent The Strings in this List will be added to the new row in the same order.
     *                   If an empty List or null is passed in, the new row will remain empty.
     * @throws NullPointerException Calling this method before any sheet has been added will result in a NullPointerException,
     *                              which will remain uncaught.
     */

    public void addRowToCurrentSheet(List<String> rowContent) {
        XSSFRow currentSheetRow = currentSheet.createRow(rowIndex++);

        if (rowContent == null) {
            return;
        }

        currentColumnAmount = Math.max(rowContent.size(), currentColumnAmount);

        for (int index = 0; index < rowContent.size(); index++) {
            XSSFCell currentCell = currentSheetRow.createCell(index);
            currentCell.setCellValue(rowContent.get(index));
            currentCell.setCellStyle(currentCellStyle);
        }
    }

    /**
     * Adds several new rows at the bottom of the current sheet.
     * <p>
     * This is a convenience method, that calls
     * <code>public void addRowToCurrentSheet(List&lt;String&gt; rowContent)</code> once for each
     * <code>List&lt;String&gt;</code> in the main <code>List&lt;List&lt;String&gt;&gt;</code>.
     * <p>
     * This method is called by several higher level convenience methods.
     *
     * @param rowContentList The content to be te added to the current sheet.
     */

    public void addSeveralRowsToCurrentSheet(List<List<String>> rowContentList) {
        for (List<String> rowContent : rowContentList) {
            addRowToCurrentSheet(rowContent);
        }
    }

    /**
     * Add a new sheet to the .xlsx file.
     * <p>
     * This is a basic method, that is called by several other methods.
     *
     * @param sheetName The name of the sheet to be added to the .xlsx file.
     *                  If null or a blank String is passed in, the sheet will be given a name.
     *                  If the passed in String consists of more than 31 characters, only the first 31 characters will be used.
     *                  If a sheet with the same name is already present in the in memory workbook, a unique name will be provided.
     */

    public void addSheetToWorkbook(String sheetName) {
        SkinnyUtil.adjustColumnSizesInCurrentSheet(currentSheet, currentColumnAmount);
        createNewSheet(sheetName);
    }

    /**
     * This method adds a sheet to the in memory workbook.
     * <p>
     * This is a convenience method that calls several basic methods of this class.
     * <p>
     * This method calls at least three out of four methods in the SkinnySheetContent interface.
     * The interface method <code>getColumnHeaders()</code> will only be called if <code>hasColumnHeaders()</code> returns true.
     * If, in that case, <code>getColumnHeaders()</code> returns null, an empty List, a List containing null, or a List
     * containing any blank String, an IllegalArgumentException or a NullPointerException will be thrown by the
     * <code>addColumnHeaderRowToCurrentSheet()</code> method and will remain uncaught.
     *
     * @param sheetContent Any implementation of this interface should provide the name of the sheet to be added, whether column
     *                     headers should be added to the sheet, the values for the column headers (if applicable),
     *                     and the content rows to be added to the sheet.
     */

    public void addSheetToWorkbook(SkinnySheetContent sheetContent) {
        addSheetToWorkbook(sheetContent.getSheetName());
        if (sheetContent.hasColumnHeaders()) {
            addColumnHeaderRowToCurrentSheet(sheetContent.getColumnHeaders());
        }
        addSeveralRowsToCurrentSheet(sheetContent.getContentRows());
    }

    /**
     * This method adds several sheets to the in memory workbook.
     * <p>
     * This is a convenience method that calls <code>addSheetToWorkbook(SkinnySheetContent sheetContent)</code> once for each
     * entry in the <code>List&lt;SkinnySheetContent&gt;</code> that is passed in as a parameter.
     *
     * @param sheetList A List of objects implementing the SkinnySheetContent interface.
     *                  Each object in the List represents a sheet to be added to the .xlsx file.
     */
    public void addSeveralSheetsToWorkbook(List<SkinnySheetContent> sheetList) {
        for (SkinnySheetContent sheetContent : sheetList) {
            addSheetToWorkbook(sheetContent);
        }

    }

    /**
     * Adds a single sheet with no column header row and zero or more content rows.
     * <p>
     * This is a convenience method that calls several basic methods of this class.
     *
     * @param sheetName    The name of the sheet to be added to the .xlsx file.
     *                     If null or an empty String is passed in, the sheet will be given a name.
     *                     If a non-unique sheet name is passed in, the sheet will be given a unique name.
     * @param sheetContent The content to be te added to the new sheet.
     *                     Null values are allowed and wil result in an empty sheet.
     */

    public void addSheetWithContentToWorkbook(String sheetName, List<List<String>> sheetContent) {
        addSheetToWorkbook(sheetName);
        if (sheetContent != null) {
            addSeveralRowsToCurrentSheet(sheetContent);
        }
    }

    /**
     * Adds a single sheet with a single column header row and zero or more content rows.
     * <p>
     * This is a convenience method that calls several basic methods of this class.
     *
     * @param sheetName              The name of the sheet to be added to the .xlsx file.
     *                               If null or an empty String is passed in, the sheet will be given a name.
     *                               If a non-unique sheet name is passed in, the sheet will be given a unique name.
     * @param sheetHeadersAndContent The first List&lt;String&gt; entry will be added as a column header row to the new sheet.
     *                               Every entry after the first (if any) will be added as a content row.
     */

    public void addSheetWithHeadersAndContentToWorkbook(String sheetName, List<List<String>> sheetHeadersAndContent) {
        addSheetToWorkbook(sheetName);
        addColumnHeaderRowToCurrentSheet(sheetHeadersAndContent.get(0));
        for (int index = 1; index < sheetHeadersAndContent.size(); index++) {
            addRowToCurrentSheet(sheetHeadersAndContent.get(index));
        }
    }

    /**
     * Adds several sheets with no column header row and zero or more content rows.
     * <p>
     * This is a convenience method that calls
     * <code>addSheetWithContentToWorkbook(String sheetName, List&lt;List&lt;String&gt;&gt; sheetContent)</code>
     * once for each entry in the Map.
     * <p>
     * If the order of the sheets is relevant, using a LinkedHashMap or similar Map implementation is highly recommended.
     *
     * @param sheetNameAndContentMap For each entry in this map, a sheet will be added to the .xlsx file.
     *                               The key of each entry will be the name of the new sheet.
     *                               The value of each entry (if any) will be the content of the corresponding sheet.
     */

    public void addSeveralSheetsWithContentToWorkbook(Map<String, List<List<String>>> sheetNameAndContentMap) {
        sheetNameAndContentMap.forEach(this::addSheetWithContentToWorkbook);
    }

    /**
     * Adds several sheets with a single column header row and zero or more content rows.
     * <p>
     * This is a convenience method that calls
     * <code>addSheetWithHeadersAndContentToWorkbook(String sheetName, List&lt;List&lt;String&gt;&gt; sheetContent)</code>
     * once for each entry in the Map.
     * <p>
     * If the order of the sheets is relevant, using a LinkedHashMap or similar Map implementation is highly recommended.
     *
     * @param sheetNameAndHeadersAndContentMap For each entry in this map, a sheet will be added to the .xlsx file.
     *                                         The key of each entry will be the name of the new sheet.
     *                                         The value of each entry will be the content of the corresponding sheet, with the first <code>List&lt;String&gt;</code> added
     *                                         as a column header row to the top of the sheet, and any subsequent <code>List&lt;String&gt;</code> (if any) as content rows.
     */

    public void addSeveralSheetsWithHeadersAndContentToWorkbook(Map<String, List<List<String>>> sheetNameAndHeadersAndContentMap) {
        sheetNameAndHeadersAndContentMap.forEach(this::addSheetWithHeadersAndContentToWorkbook);
    }

    /**
     * Writes a new .xlsx file on the file system with all added sheets and rows, writing over any previous version.
     * <p>
     * This is a basic method, that is called by the constructor when writing an empty .xlsx file with a single sheet.
     *
     * @throws IOException Any Exception that occurs while creating a file on the file system or writing to this file
     *                     will remain uncaught.
     */

    public void writeToFile() throws IOException {
        SkinnyUtil.adjustColumnSizesInCurrentSheet(currentSheet, currentColumnAmount);
        targetFile.createNewFile();
        FileOutputStream outputStream = new FileOutputStream(targetFile);
        workbook.write(outputStream);
        outputStream.close();
    }

    /**
     * Returns a String representation, including information considered useful for debugging or logging.
     *
     * @return A String representation, containing the target .xlsx file, the current amount of sheets, and the amount of rows and
     * columns on the current sheet.
     */

    @Override
    public String toString() {
        return String.format("SkinnyWriter - target .xlsx file: %s - current amount of sheets: %s - current sheet has %s rows and "
                        + "%s columns", targetFile.toString(), workbook.getNumberOfSheets(),
                currentSheet.getPhysicalNumberOfRows(), currentColumnAmount);
    }

    private void createNewSheet(String sheetName) {
        currentSheet = workbook.createSheet(SkinnyUtil.sanitizeSheetName(sheetName, workbook));
        currentCellStyle = workbook.createCellStyle();
        currentCellStyle.setWrapText(false);
        currentColumnAmount = 0;
        rowIndex = 0;
    }

}
