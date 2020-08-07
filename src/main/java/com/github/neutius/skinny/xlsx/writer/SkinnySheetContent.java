package com.github.neutius.skinny.xlsx.writer;

import java.util.List;

public interface SkinnySheetContent {

    String getSheetName();

    boolean hasColumnHeaders();

    List<String> getColumnHeaders();

    List<List<String>> getContentRows();

}
