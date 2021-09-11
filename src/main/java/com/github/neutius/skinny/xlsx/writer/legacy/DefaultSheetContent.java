package com.github.neutius.skinny.xlsx.writer.legacy;

import java.util.List;

/**
 * A default implementation for the <code>SkinnySheetContent</code> interface - see the JavaDoc for that interface for more
 * information.
 * <p>
 * This class has a private constructor and two static factory methods.
 */

public final class DefaultSheetContent implements SkinnySheetContent {

    private static final String EXCEPTION_NULL_LIST = "List of column headers should not be null";
    private static final String EXCEPTION_EMPTY_LIST = "List of column headers should not be empty";
    private static final String EXCEPTION_LIST_CONTAINS_INVALID_VALUE
            = "List of column headers should only contain String values with at least 1 non-whitespace character";

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
     *                      Cannot be null or empty, and can only contain String values with at least 1 non-whitespace character.
     * @param contentRows   Represents the content rows: zero or more Lists
     *                      containing zero or more String values to be added as content cell values.
     * @return A representation of a sheet to be added to a .xlsx file by the SkinnyWriter class.
     * Implements the SkinnySheetContent interface, allowing this return value to be passed directly to the SkinnyWriter class.
     *
     * @throws IllegalArgumentException An Exception will be thrown if the <code>List&lt;String&gt; columnHeaders</code>
     * is null, is empty, contains any null value, or contains any blank String.
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

    private DefaultSheetContent(String sheetName, boolean hasColumnHeaders, List<String> columnHeaders,
                                List<List<String>> contentRows) {
        this.sheetName = sheetName;
        this.hasColumnHeaders = hasColumnHeaders;
        this.columnHeaders = sanitizeColumnHeaders(hasColumnHeaders, columnHeaders);
        this.contentRows = contentRows;
    }

    private List<String> sanitizeColumnHeaders(boolean hasColumnHeaders, List<String> columnHeaders) {
        if (!hasColumnHeaders) {
            return null;
        }
        throwExceptionIfParameterIsInvalid(columnHeaders);
        return columnHeaders;
    }

    private void throwExceptionIfParameterIsInvalid(List<String> columnHeaders) {
        if (columnHeaders == null) {
            throw new IllegalArgumentException(EXCEPTION_NULL_LIST);
        }
        if (columnHeaders.isEmpty()) {
            throw new IllegalArgumentException(EXCEPTION_EMPTY_LIST);
        }
        if (listContainsNullOrBlankString(columnHeaders)) {
            throw new IllegalArgumentException(EXCEPTION_LIST_CONTAINS_INVALID_VALUE);
        }
    }

    private boolean listContainsNullOrBlankString(List<String> stringList) {
        return stringList.stream().anyMatch(value -> value == null || value.isBlank());
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
