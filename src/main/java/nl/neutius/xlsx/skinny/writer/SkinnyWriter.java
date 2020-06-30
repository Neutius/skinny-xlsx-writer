package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class SkinnyWriter {

    private static final String EXTENSION = ".xlsx";

    private final File targetFile;

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
     * Calling this constructor will initialize an in memory Workbook with a single empty Sheet, which will be
     * immediately written to disk. This ensures any IOException will be encountered as quickly as possible.
     *
     * Warning: if the target directory already has a .xlsx file with the same name,
     * it will be overwritten with no further warning.
     *
     * @param targetFolder The target directory where the .xlsx file will be written to. Must be an existing directory.
     * @param fileName The base name of the .xlsx that will be written.
     *                 No extension needed, this class will add the .xlsx for you.
     *                 If null or an empty String is passed in, the file will be given a name.
     * @param firstSheetName The name of the first sheet of the .xlsx file.
     *                       If null or an empty String is passed in, the sheet will be given a name.
     * @throws IOException Any Exception that occurs while creating a file on the file system or writing to this file
     * will remain uncaught.
     */

    public SkinnyWriter(File targetFolder, String fileName, String firstSheetName) throws IOException {
        targetFile = new File(targetFolder, sanitizeFileName(fileName) + EXTENSION);
        workbook = new XSSFWorkbook();
        createNewSheet(firstSheetName);
        writeToFile();
    }

    /**
     * Writes a new .xlsx file on the file system with all added sheets and rows, writing over any previous version.
     * @throws IOException Any Exception that occurs while creating a file on the file system or writing to this file
     * will remain uncaught.
     */

    public void writeToFile() throws IOException {
        adjustColumnSizesInCurrentSheet();
        targetFile.createNewFile();
        FileOutputStream outputStream = new FileOutputStream(targetFile);
        workbook.write(outputStream);
        outputStream.close();
    }

    /**
     * Convenience method to a sheet with headers and content to the .xlsx file.  Calls several basic method of this class.
     * @param sheetName The name of the sheet to be added to the .xlsx file.
     *                  If null or an empty String is passed in, the sheet will be given a name.
     * @param sheetContent The content to be te added to the new sheet.
     */

    public void addSheetWithContentToWorkbook(String sheetName, List<List<String>> sheetContent) {
        addSheetToWorkbook(sheetName);
        addSeveralRowsToCurrentSheet(sheetContent);
    }

    /**
     * Convenience method that call @addSheetWithContentToWork once for each entry.
     * @param sheetNameAndContentMap For each entry in this map, a sheet will be added to the .xlsx file.
     *                               The key of each entry will be the name of the new sheet.
     *                               The value of each entry will be the content of the corresponding sheet.
     */

    public void addSeveralSheetsWithContentToWorkbook(Map<String, List<List<String>>> sheetNameAndContentMap) {
        sheetNameAndContentMap.entrySet().forEach((entry) -> addSheetWithContentToWorkbook(entry.getKey(), entry.getValue()));
    }

    /**
     * Convenience method to a sheet with headers and content to the .xlsx file. Calls several basic method of this class.
     * @param sheetName The name of the sheet to be added to the .xlsx file.
     *                  If null or an empty String is passed in, the sheet will be given a name.
     * @param sheetHeadersAndContent The first List<String> entry will be added as a column header row to the new sheet,
     *                               every entry after the first will be added as a content row.
     */

    public void addSheetWithHeadersAndContentToWorkbook(String sheetName, List<List<String>> sheetHeadersAndContent) {
        addSheetToWorkbook(sheetName);
        addColumnHeaderRowToCurrentSheet(sheetHeadersAndContent.get(0));
        for (int index = 1; index < sheetHeadersAndContent.size(); index++) {
            addRowToCurrentSheet(sheetHeadersAndContent.get(index));
        }
    }

    /**
     * Add a new sheet to the .xlsx file.
     *
     * @param sheetName The name of the sheet to be added to the .xlsx file.
     *                  If null or an empty String is passed in, the sheet will be given a name.
     */

    public void addSheetToWorkbook(String sheetName) {
        adjustColumnSizesInCurrentSheet();
        createNewSheet(sheetName);
    }

    /**
     * Adds a row at the top of the current sheet, and adds the parameter values to this row. This row is unique in three ways:
     * 1. Column header text will be given a bold font.
     * 2. A freeze pane will be applied to the column header row.
     * 3. Line wrapping is disabled for column header cells.
     *
     * Adding column headers to a sheet is optional.
     *
     * Column headers should be added first and only once:
     *      Any attempt to add column headers to a non-empty sheet will result in an IllegalStateException.
     *
     * Column header text should consist of at least 1 non white space character:
     *      If the parameter contains any null values, a NullPointerException will occur and remain uncaught.
     *      If the parameter contains any blank Strings, an IllegalArgumentException will be thrown.
     *
     * @param columnHeaderRow The List of String values to be added to the column header row.
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

            currentCell.setCellValue(applyColumnHeaderFont(columnHeaderRow.get(columnIndex)));
        }

        currentSheet.createFreezePane(0, 1);
    }

    /**
     * Creates a new row at the bottom of the current sheet.
     * @param rowContent The Strings in this List will be added to the new row in the same order.
     *                   If an empty List or null is passed in, the new row will remain empty.
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
     * Calls <code> public void addRowToCurrentSheet(List<String> rowContent) </code> once for each
     * <code> List<String> </code> in the main <code> List<List<String>> </code>.
     * @param rowContentList The content to be te added to the current sheet.
     */

    public void addSeveralRowsToCurrentSheet(List<List<String>> rowContentList) {
        for (List<String> rowContent : rowContentList) {
            addRowToCurrentSheet(rowContent);
        }
    }

    /**
     * Only information considered useful for debugging or logging is included.
     * @return A String representation, containing the target .xlsx file, the current amount of sheets, and the amount of rows and
     * columns on the current sheet.
     */

    @Override
    public String toString() {
        return String.format("SkinnyWriter - target .xlsx file: %s - current amount of sheets: %s - current sheet has %s rows and "
                + "%s columns", targetFile, workbook.getNumberOfSheets(), currentSheet.getPhysicalNumberOfRows(), currentColumnAmount);
    }

    private void createNewSheet(String sheetName) {
        currentSheet = workbook.createSheet(sanitizeSheetName(sheetName));
        currentCellStyle = workbook.createCellStyle();
        currentCellStyle.setWrapText(false);
        currentColumnAmount = 0;
        rowIndex = 0;
    }

    private String sanitizeFileName(String fileName) {
        if (fileName == null || fileName.isBlank()) {
            return "output-at-" + new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss").format(new Date());
        }
        return fileName;
    }

    private String sanitizeSheetName(String sheetName) {
        if (sheetName == null || sheetName.isBlank()) {
            return "Sheet_" + (workbook.getNumberOfSheets() + 1);
        }
        if (workbook.getSheet(sheetName) != null) {
            return sheetName + '_' + (workbook.getNumberOfSheets() + 1);
        }

        return sheetName;
    }

    private void adjustColumnSizesInCurrentSheet() {
        currentCellStyle.setWrapText(false);

        for (int index = 0; index < currentColumnAmount; index++) {
            currentSheet.autoSizeColumn(index);
        }

        currentCellStyle.setWrapText(true);
    }

    private XSSFRichTextString applyColumnHeaderFont(String originalValue) {
        XSSFRichTextString richValue = new XSSFRichTextString(originalValue);
        richValue.applyFont(columnHeaderFont);
        return richValue;
    }
}
