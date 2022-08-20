package com.github.neutius.skinny.xlsx.writer;

import com.github.neutius.skinny.xlsx.writer.interfaces.ContentRowSupplier;
import com.github.neutius.skinny.xlsx.writer.interfaces.SheetContentSupplier;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

public class SkinnySheetContentSupplier implements SheetContentSupplier {
	private final List<ContentRowSupplier> contentRowSuppliers = new ArrayList<>();

	@Override
	public List<ContentRowSupplier> get() {
		return contentRowSuppliers;
	}

	public SkinnySheetContentSupplier() {
	}

	public SkinnySheetContentSupplier(Collection<ContentRowSupplier> initialContent) {
		initialContent.forEach(this::addContentRowSupplierToSheet);
	}

	public SkinnySheetContentSupplier(ContentRowSupplier... initialContent) {
		Arrays.asList(initialContent).forEach(this::addContentRowSupplierToSheet);
	}

	public void addContentRowSupplier(ContentRowSupplier contentRowSupplier) {
		addContentRowSupplierToSheet(contentRowSupplier);
	}

	private void addContentRowSupplierToSheet(ContentRowSupplier contentRowSupplier) {
		if (contentRowSupplier == null || contentRowSupplier.get() == null) {
			contentRowSuppliers.add(new SkinnyContentRowSupplier());
		} else {
			contentRowSuppliers.add(contentRowSupplier);
		}
	}

	public void addContentRow(Collection<String> contentRow) {
		if (contentRow == null) {
			contentRowSuppliers.add(new SkinnyContentRowSupplier());
		} else {
			contentRowSuppliers.add(new SkinnyContentRowSupplier(contentRow));
		}
	}

	public void addContentRow(String... contentRow) {
		contentRowSuppliers.add(new SkinnyContentRowSupplier(contentRow));
	}

}
