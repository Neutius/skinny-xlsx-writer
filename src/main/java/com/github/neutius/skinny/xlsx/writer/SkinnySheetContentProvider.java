package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxRowContentProvider;
import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxSheetContentProvider;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class SkinnySheetContentProvider implements XlsxSheetContentProvider {
    private final ArrayList<XlsxRowContentProvider> rowContentProviders;

    // TODO Constructor overloads to feed RowContentProvider as Collection or as var args - GvdNL 26-09-2021
    public SkinnySheetContentProvider() {
        rowContentProviders = new ArrayList<>();
    }

    @Override
    public List<XlsxRowContentProvider> getRowContentProviders() {
        return rowContentProviders;
    }

    public void addRowContentProvider(XlsxRowContentProvider rowContentProvider) {
        rowContentProviders.add(rowContentProvider);
    }

    public void addContentRow(String... rowContent) {
        rowContentProviders.add(new SkinnyRowContentProvider(rowContent));
    }

    public void addContentRow(Collection<String> rowContent) {
        rowContentProviders.add(new SkinnyRowContentProvider(rowContent));
    }

}