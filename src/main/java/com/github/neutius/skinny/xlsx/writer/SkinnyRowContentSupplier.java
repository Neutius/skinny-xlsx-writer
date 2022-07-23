package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkinnyRowContentSupplier implements RowContentSupplier {
    private final List<String> rowContent;

    @Override
    public List<String> get() {
        return rowContent;
    }

    public SkinnyRowContentSupplier() {
        rowContent = new ArrayList<>();
    }

    public SkinnyRowContentSupplier(Collection<String> initialContent) {
        rowContent = new ArrayList<>(initialContent);
    }

    public SkinnyRowContentSupplier(String... initialContent) {
        rowContent = new ArrayList<>(Arrays.asList(initialContent));
    }

    public void addCellContent(String content) {
        rowContent.add(sanitizeCellContent(content));
    }

    private static String sanitizeCellContent(String content) {
        if (content == null || content.isBlank()) {
            return "";
        }
        return content;
    }

}
