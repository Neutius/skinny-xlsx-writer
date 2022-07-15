package com.github.neutius.skinny.xlsx.writer.interfaces;

import java.util.List;
import java.util.function.Supplier;

/*
 Is this a useful extension of a standard functional interface? Should the method declaration be there?
 Perhaps it's better to use this as an empty marker interface, with JavaDoc? - GvdNL 15-07-2022
 */

public interface RowContentSupplier extends Supplier<List<String>> {

    List<String> get();

}
