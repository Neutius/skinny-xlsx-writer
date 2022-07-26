package com.github.neutius.skinny.xlsx.writer.interfaces;

import java.util.List;
import java.util.function.Supplier;

/*
 Is this a useful extension of a standard functional interface? Should the method declaration be there?
 Perhaps it's better to use this as an empty marker interface, with JavaDoc? - GvdNL 15-07-2022
 */

/*
 A sheet has more than content: it also has a name and optionally also column headers.

 Adding a new functional interface "SheetSupplier" wouldn't work: a Sheet is created by the workbook.
 The data for the sheet (name, content rows, column headers) must be held by some object.
 This can still be defined by an interface - that would be preferable - but one with several methods.

 Adding a "SheetProvider" interface with "getSheetContentProvider", "getColumnHeaders" and "getSheetName" methods might be best.
 Any implementation may create or consume a SheetContentSupplier instance.

 Perhaps add a SheetNameSupplier and/or a ColumnHeaderSupplier?
 SheetNameSupplier might be overkill for a single String - or would it be useful for client code to inject their own name generator?

  - GvdNL 15-07-2022
 */


public interface SheetContentSupplier extends Supplier<List<ContentRowSupplier>> {

	List<ContentRowSupplier> get();

}
