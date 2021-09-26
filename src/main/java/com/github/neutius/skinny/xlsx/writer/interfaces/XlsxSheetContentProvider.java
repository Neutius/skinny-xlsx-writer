package com.github.neutius.skinny.xlsx.writer.interfaces;

import java.util.List;

public interface XlsxSheetContentProvider {

    List<XlsxRowContentProvider> getRowContentProviders();

}