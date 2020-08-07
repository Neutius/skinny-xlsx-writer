package com.github.neutius.skinny.xlsx.writer;

import java.util.List;

public final class DefaultSheetContent implements SkinnySheetContent {

    private String sheetName;
    private boolean hasColumnHeaders;
    private List<String> columnHeaders;
    private List<List<String>> contentRows;

    public static DefaultSheetContent withHeaders(String sheetName, List<String> columnHeaders, List<List<String>> contentRows) {
        return new DefaultSheetContent(sheetName, true, columnHeaders, contentRows);
    }

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
