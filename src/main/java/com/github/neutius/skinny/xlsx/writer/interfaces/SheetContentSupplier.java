package com.github.neutius.skinny.xlsx.writer.interfaces;

import java.util.List;
import java.util.function.Supplier;

/*
 Is this a useful extension of a standard functional interface? Should the method declaration be there?
 Perhaps it's better to use this as an empty marker interface, with JavaDoc?
 Or perhaps remove it? - GvdNL 15-07-2022
 */

/*
 A sheet has more than content: it also has a name and optionally also column headers.

 Perhaps add a SheetNameSupplier and a ColumnHeaderProvider, both of which can be used by an implementation of SheetSupplier?
  - GvdNL 15-07-2022
 */


public interface SheetContentSupplier extends Supplier<List<RowContentSupplier>> {

	List<RowContentSupplier> get();

}
