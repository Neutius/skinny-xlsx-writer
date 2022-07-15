package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class SkinnySheetContentSupplier implements SheetContentSupplier {
    private final ArrayList<RowContentSupplier> rowContentProviders;

    // TODO Constructor overloads to feed RowContentProvider as Collection or as var args - GvdNL 26-09-2021
    public SkinnySheetContentSupplier() {
        rowContentProviders = new ArrayList<>();
    }

    @Override
    public List<RowContentSupplier> get() {
        return rowContentProviders;
    }

    public void addRowContentProvider(RowContentSupplier rowContentProvider) {
        rowContentProviders.add(rowContentProvider);
    }

    public void addContentRow(String... rowContent) {
        rowContentProviders.add(new SkinnyRowContentSupplier(rowContent));
    }

    public void addContentRow(Collection<String> rowContent) {
        rowContentProviders.add(new SkinnyRowContentSupplier(rowContent));
    }

}