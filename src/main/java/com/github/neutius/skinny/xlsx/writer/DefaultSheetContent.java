package com.github.neutius.skinny.xlsx.writer;

import java.util.List;

/**
 * A default implementation for the <code>SkinnySheetContent</code> interface - see the JavaDoc for that interface for more
 * information.
 * <p>
 * This class has a private constructor and two static factory methods.
 */

public final class DefaultSheetContent implements SkinnySheetContent {

    private final String sheetName;
    private final boolean hasColumnHeaders;
    private final List<String> columnHeaders;
    private final List<List<String>> contentRows;

    /**
     * This method creates and returns a representation of a sheet to be added to a .xlsx file by the SkinnyWriter class,
     * with a single column header row at the top of the sheet.
     *
     * @param sheetName     The name of the sheet to be added.
     * @param columnHeaders Represents the column header row: A List of String values to be added to the sheet as column headers.
     * @param contentRows   Represents the content rows: zero or more Lists
     *                      containing zero or more String values to be added as content cell values.
     * @return A representation of a sheet to be added to a .xlsx file by the SkinnyWriter class.
     * Implements the SkinnySheetContent interface, allowing this return value to be passed directly to the SkinnyWriter class.
     */

    public static DefaultSheetContent withHeaders(String sheetName, List<String> columnHeaders, List<List<String>> contentRows) {
        return new DefaultSheetContent(sheetName, true, columnHeaders, contentRows);
    }

    /**
     * This method creates and returns a representation of a sheet to be added to a .xlsx file by the SkinnyWriter class,
     * with no column header row.
     *
     * @param sheetName     The name of the sheet to be added.
     * @param contentRows   Represents the content rows: zero or more Lists
     *                      containing zero or more String values to be added as content cell values.
     * @return A representation of a sheet to be added to a .xlsx file by the SkinnyWriter class.
     * Implements the SkinnySheetContent interface, allowing this return value to be passed directly to the SkinnyWriter class.
     */

    public static DefaultSheetContent withoutHeaders(String sheetName, List<List<String>> contentRows) {
        return new DefaultSheetContent(sheetName, false, null, contentRows);
    }

    private DefaultSheetContent(String sheetName, boolean hasColumnHeaders, List<String> columnHeaders, List<List<String>> contentRows) {
        this.sheetName = sheetName;
        this.hasColumnHeaders = hasColumnHeaders;
        this.columnHeaders = hasColumnHeaders ? columnHeaders : null;
        this.contentRows = contentRows;
    }

    @Override
    public String getSheetName() {
        return sheetName;
    }

    @Override
    public boolean hasColumnHeaders() {
        return hasColumnHeaders;
    }

    @Override
    public List<String> getColumnHeaders() {
        return columnHeaders;
    }

    @Override
    public List<List<String>> getContentRows() {
        return contentRows;
    }

}
