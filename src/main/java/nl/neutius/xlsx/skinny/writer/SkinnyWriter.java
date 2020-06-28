package nl.neutius.xlsx.skinny.writer;

import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class SkinnyWriter {

    private static final String EXTENSION = ".xlsx";

    private final File targetFile;
    private final String firstSheetName;

    private XSSFWorkbook workbook;
    private XSSFCellStyle currentCellStyle;
    private XSSFSheet currentSheet;
    private int currentColumnAmount;
    private int rowIndex;

    public SkinnyWriter(File targetFolder, String fileName, String firstSheetName) {
        targetFile = new File(targetFolder, fileName + EXTENSION);
        this.firstSheetName = firstSheetName;
    }

    public void createNewXlsxFile() throws IOException {
        workbook = new XSSFWorkbook();
        createNewSheet(firstSheetName);
        writeToFile();
    }

    public void addSheetToWorkbook(String sheetName) {
        adjustCellSizesInCurrentSheet();
        createNewSheet(sheetName);
    }

    private void createNewSheet(String sheetName) {
        currentSheet = workbook.createSheet(sanitizeSheetName(sheetName));
        currentCellStyle = workbook.createCellStyle();
        currentCellStyle.setWrapText(false);
        currentColumnAmount = 0;
        rowIndex = 0;
    }

    private String sanitizeSheetName(String sheetName) {
        if (sheetName == null ||sheetName.isEmpty()) {
            int newSheetNumber = workbook.getNumberOfSheets() + 1;
            return "Sheet_" + newSheetNumber;
        }
        return sheetName;
    }

    public void addRowToCurrentSheet(List<String> rowContent) {
        XSSFRow currentSheetRow = currentSheet.createRow(rowIndex++);

        currentColumnAmount = Math.max(rowContent.size(), currentColumnAmount);

        for (int index = 0; index < rowContent.size(); index++) {
            XSSFCell currentCell = currentSheetRow.createCell(index);
            currentCell.setCellValue(rowContent.get(index));
            currentCell.setCellStyle(currentCellStyle);
        }
    }

    public void addSeveralRowsToCurrentSheet(List<List<String>> rowContentList) {
        for (List<String> rowContent : rowContentList) {
            addRowToCurrentSheet(rowContent);
        }
    }

    public void writeToFile() throws IOException {
        adjustCellSizesInCurrentSheet();
        targetFile.createNewFile();
        FileOutputStream outputStream = new FileOutputStream(targetFile);
        workbook.write(outputStream);
        outputStream.close();
    }

    private void adjustCellSizesInCurrentSheet() {
        currentCellStyle.setWrapText(false);

        for (int index = 0; index < currentColumnAmount; index++) {
            currentSheet.autoSizeColumn(index);
        }

        currentCellStyle.setWrapText(true);
    }
}
