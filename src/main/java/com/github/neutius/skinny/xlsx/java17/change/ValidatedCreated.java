package com.github.neutius.skinny.xlsx.java17.change;

public class ValidatedCreated extends Created {

	public ValidatedCreated(Object initialObject, Object newObject) {
		super(initialObject, newObject);
		if (initialObject != null) {
			throw new IllegalArgumentException("For change of type 'created': initialObject should be null");
		}
		if (newObject == null) {
			throw new IllegalArgumentException("For change of type 'created': newObject cannot be null");
		}
	}

}
