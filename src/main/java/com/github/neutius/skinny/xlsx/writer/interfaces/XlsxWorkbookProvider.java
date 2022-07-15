package com.github.neutius.skinny.xlsx.writer.interfaces;

import org.apache.poi.ss.usermodel.Workbook;

// Perhaps annotate this as a functional interface? - GvdNL 15-07-2022

// Why return a workbook? Why not a List<Sheet> or a List<SheetSupplier>? - GvdNL 15-07-2022

public interface XlsxWorkbookProvider {

    Workbook getWorkbook();

}
