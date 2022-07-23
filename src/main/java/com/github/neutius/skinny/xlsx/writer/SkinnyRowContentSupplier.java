package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkinnyRowContentSupplier implements RowContentSupplier {
    private final List<String> rowContent = new ArrayList<>();

    @Override
    public List<String> get() {
        return rowContent;
    }

    public SkinnyRowContentSupplier() {
    }

    public SkinnyRowContentSupplier(Collection<String> initialContent) {
        if (initialContent != null) {
            initialContent.forEach(this::addCellContentToRow);
        }
    }

    public SkinnyRowContentSupplier(String... initialContent) {
        Arrays.asList(initialContent).forEach(this::addCellContentToRow);
    }

    public void addCellContent(String content) {
        addCellContentToRow(content);
    }

    private void addCellContentToRow(String content) {
        rowContent.add(sanitizeCellContent(content));
    }

    private static String sanitizeCellContent(String content) {
        if (content == null || content.isBlank()) {
            return "";
        }
        return content;
    }

}
