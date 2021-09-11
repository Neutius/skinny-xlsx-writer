package com.github.neutius.skinny.xlsx.writer.legacy;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * This class has a single public static method that writes a .xlsx file to disk
 * <p>
 * This class uses the Apache POI SXSSF streaming API to improve performance.
 * <p>
 * This class is currently in beta.
 */

public final class SkinnyStreamer {
    private final File targetFile;
    private final SXSSFWorkbook workbook;
    private final CellStyle columnHeaderCellStyle;

    private int currentColumnAmount = 1;

    /**
     * Offers basically the same functionality as the SkinnyWriter method of the same name - there might be some small differences.
     * <p>
     * This method uses the Apache POI SXSSF streaming API to improve performance.
     * <p>
     * This method is currently in beta.
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

    public static void writeContentToFileSystem(File targetFolder, String fileName, List<SkinnySheetContent> sheetContentList) throws IOException {
        SkinnyStreamer streamer = new SkinnyStreamer(targetFolder, fileName);
        streamer.addSeveralSheetsToWorkbook(sheetContentList);
        streamer.writeToFile();
        streamer.cleanUp();
    }

    private SkinnyStreamer(File targetFolder, String fileName) {
        targetFile = new File(targetFolder, SkinnyUtil.sanitizeFileName(fileName) + SkinnyUtil.EXTENSION);
        workbook = new SXSSFWorkbook();
        columnHeaderCellStyle = SkinnyUtil.createColumnHeaderCellStyle(workbook);
    }

    private void addSeveralSheetsToWorkbook(List<SkinnySheetContent> sheetContentList) {
        for (SkinnySheetContent content : sheetContentList) {
                addSheetToWorkbook(content);
        }
    }

    private void addSheetToWorkbook(SkinnySheetContent content) {
        SXSSFSheet currentSheet = workbook.createSheet(SkinnyUtil.sanitizeSheetName(content.getSheetName(), workbook));
        if (content.hasColumnHeaders()) {
            addColumnHeaderRow(currentSheet, content.getColumnHeaders());
        }
        addContentRows(currentSheet, content.getContentRows());
    }

    private void addColumnHeaderRow(SXSSFSheet currentSheet, List<String> columnHeaders) {
        SXSSFRow headerRow = currentSheet.createRow(currentSheet.getPhysicalNumberOfRows());

        for (String text : columnHeaders) {
            SXSSFCell cell = headerRow.createCell(headerRow.getPhysicalNumberOfCells());
            cell.setCellValue(text);
            cell.setCellStyle(columnHeaderCellStyle);
        }

        keepTrackOfColumnAmount(columnHeaders);
        currentSheet.createFreezePane(0, 1);
    }

    private void addContentRows(SXSSFSheet currentSheet, List<List<String>> contentRows) {
        for (List<String> contentRow : contentRows) {
            addContentRow(currentSheet, contentRow);
        }

        if (currentSheet.getPhysicalNumberOfRows() < 100) {
            adjustColumnSizes(currentSheet);
        }
    }

    private void addContentRow(SXSSFSheet currentSheet, List<String> contentRow) {
        if (contentRow == null) {
            currentSheet.createRow(currentSheet.getPhysicalNumberOfRows());
            return;
        }

        SXSSFRow row = currentSheet.createRow(currentSheet.getPhysicalNumberOfRows());

        for (String text : contentRow) {
            SXSSFCell cell = row.createCell(row.getPhysicalNumberOfCells());
            cell.setCellValue(text);
        }

        keepTrackOfColumnAmount(contentRow);

        if (currentSheet.getPhysicalNumberOfRows() == 100) {
            adjustColumnSizes(currentSheet);
        }

    }

    private void keepTrackOfColumnAmount(List<String> rowAdded) {
        currentColumnAmount = Math.max(rowAdded.size(), currentColumnAmount);
    }

    private void adjustColumnSizes(SXSSFSheet currentSheet) {
        currentSheet.trackAllColumnsForAutoSizing();
        SkinnyUtil.adjustColumnSizesInCurrentSheet(currentSheet, currentColumnAmount);
        currentSheet.untrackAllColumnsForAutoSizing();
    }

    private void writeToFile() throws IOException {
        targetFile.createNewFile();
        FileOutputStream outputStream = new FileOutputStream(targetFile);
        workbook.write(outputStream);
        outputStream.close();
    }

    // Note that SXSSF allocates temporary files that you must always clean up explicitly, by calling the dispose method.
    private void cleanUp() {
        workbook.dispose();
    }

}
