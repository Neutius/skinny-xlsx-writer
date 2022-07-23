package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkinnySheetContentSupplier implements SheetContentSupplier {
    private final List<RowContentSupplier> rowContentSuppliers = new ArrayList<>();

    @Override
    public List<RowContentSupplier> get() {
        return rowContentSuppliers;
    }

    public SkinnySheetContentSupplier() {
    }

    public SkinnySheetContentSupplier(Collection<RowContentSupplier> initialContent) {
        initialContent.forEach(this::addRowContentSupplierToSheet);
    }

    public SkinnySheetContentSupplier(RowContentSupplier... initialContent) {
        Arrays.asList(initialContent).forEach(this::addRowContentSupplierToSheet);
    }

    public void addRowContentSupplier(RowContentSupplier rowContentSupplier) {
        addRowContentSupplierToSheet(rowContentSupplier);
    }

    private void addRowContentSupplierToSheet(RowContentSupplier rowContentSupplier) {
        if (rowContentSupplier == null || rowContentSupplier.get() == null) {
            rowContentSuppliers.add(new SkinnyRowContentSupplier());
        }
        else {
            rowContentSuppliers.add(rowContentSupplier);
        }
    }

    public void addContentRow(Collection<String> rowContent) {
        if (rowContent == null) {
            rowContentSuppliers.add(new SkinnyRowContentSupplier());
        }
        else {
            rowContentSuppliers.add(new SkinnyRowContentSupplier(rowContent));
        }
    }

    public void addContentRow(String... rowContent) {
        rowContentSuppliers.add(new SkinnyRowContentSupplier(rowContent));
    }

}
