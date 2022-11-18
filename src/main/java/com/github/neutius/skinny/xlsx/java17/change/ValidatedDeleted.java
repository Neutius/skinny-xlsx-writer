package com.github.neutius.skinny.xlsx.java17.change;

public non-sealed class ValidatedDeleted extends Deleted {

	public ValidatedDeleted(Object initialObject, Object newObject) {
		super(initialObject, newObject);
		if (initialObject == null) {
			throw new IllegalArgumentException("For change of type 'deleted': initialObject cannot be null");
		}
		if (newObject != null) {
			throw new IllegalArgumentException("For change of type 'deleted': newObject should be null");
		}
	}

}
