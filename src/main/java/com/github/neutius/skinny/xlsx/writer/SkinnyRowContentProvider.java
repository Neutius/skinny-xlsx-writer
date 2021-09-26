package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.XlsxRowContentProvider;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkinnyRowContentProvider implements XlsxRowContentProvider {
    private final List<String> rowContent;

    public SkinnyRowContentProvider() {
        rowContent = new ArrayList<>();
    }

    public SkinnyRowContentProvider(Collection<String> initialContent) {
        rowContent = new ArrayList<>(initialContent);
    }

    public SkinnyRowContentProvider(String... initialContent) {
        rowContent = new ArrayList<>(Arrays.asList(initialContent));
    }

    public void addCellContent(String content) {
        rowContent.add(content);
    }

    @Override
    public List<String> getRowContent() {
        return rowContent;
    }

}