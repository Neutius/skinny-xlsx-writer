package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ContentRowSupplier;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkinnyContentRowSupplier implements ContentRowSupplier {
    private final List<String> contentRow = new ArrayList<>();

    @Override
    public List<String> get() {
        return contentRow;
    }

    public SkinnyContentRowSupplier() {
    }

    public SkinnyContentRowSupplier(Collection<String> initialContent) {
        if (initialContent != null) {
            initialContent.forEach(this::addCellContentToRow);
        }
    }

    public SkinnyContentRowSupplier(String... initialContent) {
        Arrays.asList(initialContent).forEach(this::addCellContentToRow);
    }

    public void addCellContent(String content) {
        addCellContentToRow(content);
    }

    private void addCellContentToRow(String content) {
        contentRow.add(sanitizeCellContent(content));
    }

    private static String sanitizeCellContent(String content) {
        if (content == null || content.isBlank()) {
            return "";
        }
        return content;
    }

}
