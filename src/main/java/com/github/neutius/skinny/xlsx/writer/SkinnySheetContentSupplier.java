package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkinnySheetContentSupplier implements SheetContentSupplier {
    private final List<RowContentSupplier> rowContentSuppliers;

    @Override
    public List<RowContentSupplier> get() {
        return rowContentSuppliers;
    }

    public SkinnySheetContentSupplier() {
        rowContentSuppliers = new ArrayList<>();
    }

    public SkinnySheetContentSupplier(Collection<RowContentSupplier> initialContent) {
        this.rowContentSuppliers = new ArrayList<>(initialContent);
    }

    public SkinnySheetContentSupplier(RowContentSupplier... initialContent) {
        this.rowContentSuppliers = Arrays.asList(initialContent);
    }

    public void addRowContentSupplier(RowContentSupplier rowContentSupplier) {
        rowContentSuppliers.add(rowContentSupplier);
    }

    public void addContentRow(Collection<String> rowContent) {
        rowContentSuppliers.add(new SkinnyRowContentSupplier(rowContent));
    }

    public void addContentRow(String... rowContent) {
        rowContentSuppliers.add(new SkinnyRowContentSupplier(rowContent));
    }

}
