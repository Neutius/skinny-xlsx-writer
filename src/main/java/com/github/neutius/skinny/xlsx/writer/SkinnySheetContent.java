package com.github.neutius.skinny.xlsx.writer;

import java.util.List;

/**
 * This interface represents a single sheet to be added to a .xlsx file by the SkinnyWriter class.
 * <p>
 * Important note: For any instance of any implementation of this interface, if <code>hasColumnHeaders()</code> returns true, then
 * <code>getColumnHeaders()</code> has to return a List with at least one String, containing at least one non-whitespace character.
 */

public interface SkinnySheetContent {

    /**
     * This method should return the name of the sheet to be added to the .xlsx file.
     *
     * @return The name of the sheet to be added. See <code>SkinnyWriter.addSheetToWorkbook(String)</code> for handling of edge
     * cases.
     */

    String getSheetName();

    /**
     * This method informs the SkinnyWriter class whether or not column headers should be added.
     *
     * @return A boolean value, with "true" meaning "column headers should be added to the sheet", and "false" meaning the opposite.
     */

    boolean hasColumnHeaders();

    /**
     * This method should return all values for the column headers, if applicable. Each value should be a non-blank String value.
     *
     * @return A List of String values to be added, in order, as column headers.
     * If <code>hasColumnHeaders()</code> returns true, this method is not allowed to return null, an empty List, or a List
     * containing anything but String values. Each String value should contain at least one non-whitespace character.
     * If, and only if, <code>hasColumnHeaders()</code> returns false, this method will not be called,
     * and no valid return value is required.
     */

    List<String> getColumnHeaders();

    /**
     * This method should return all content values to be added to the sheet.
     *
     * @return The content values to be added to the sheet.
     * Any combination of null values, empty Lists, and blank Strings is allowed.
     */

    List<List<String>> getContentRows();

}
