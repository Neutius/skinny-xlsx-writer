package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

public class SkinnyWriter {

    private static final String EXTENSION = ".xlsx";

    private final File targetFile;

    private XSSFWorkbook workbook;
    private XSSFCellStyle currentCellStyle;
    private XSSFSheet currentSheet;
    private int currentColumnAmount;
    private int rowIndex;

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
     * will be thrown upwards.
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
     * will be thrown upwards.
     */

    public void writeToFile() throws IOException {
        adjustColumnSizesInCurrentSheet();
        targetFile.createNewFile();
        FileOutputStream outputStream = new FileOutputStream(targetFile);
        workbook.write(outputStream);
        outputStream.close();
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
}
