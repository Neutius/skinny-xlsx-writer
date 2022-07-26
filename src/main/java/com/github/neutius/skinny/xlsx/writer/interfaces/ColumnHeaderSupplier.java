package com.github.neutius.skinny.xlsx.writer.interfaces;

import java.util.List;
import java.util.function.Supplier;

public interface ColumnHeaderSupplier extends Supplier<List<String>> {

	List<String> get();

}
