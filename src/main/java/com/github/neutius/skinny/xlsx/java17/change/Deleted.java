package com.github.neutius.skinny.xlsx.java17.change;

public sealed class Deleted extends Change permits ValidatedDeleted {

	public Deleted(Object initialObject, Object newObject) {
		super(initialObject, newObject);
	}


}
