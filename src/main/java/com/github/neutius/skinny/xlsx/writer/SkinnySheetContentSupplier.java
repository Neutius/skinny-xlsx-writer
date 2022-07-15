package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class SkinnySheetContentSupplier implements SheetContentSupplier {
    private final List<RowContentSupplier> rowContentSuppliers;

    // TODO Constructor overloads to feed RowContentSupplier as Collection or as var args - GvdNL 26-09-2021
    public SkinnySheetContentSupplier() {
        rowContentSuppliers = new ArrayList<>();
    }

    @Override
    public List<RowContentSupplier> get() {
        return rowContentSuppliers;
    }

    public void addRowContentSupplier(RowContentSupplier rowContentSupplier) {
        rowContentSuppliers.add(rowContentSupplier);
    }

    public void addContentRow(String... rowContent) {
        rowContentSuppliers.add(new SkinnyRowContentSupplier(rowContent));
    }

    public void addContentRow(Collection<String> rowContent) {
        rowContentSuppliers.add(new SkinnyRowContentSupplier(rowContent));
    }

}
