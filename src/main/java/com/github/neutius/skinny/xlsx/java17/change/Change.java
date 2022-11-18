package com.github.neutius.skinny.xlsx.java17.change;

public sealed abstract class Change permits Created, Updated, Deleted {

	private final Object initialObject;
	private final Object newObject;

	public Change(Object initialObject, Object newObject) {
		this.initialObject = initialObject;
		this.newObject = newObject;
	}

	public Object getInitialObject() {
		return initialObject;
	}

	public Object getNewObject() {
		return newObject;
	}

	public static boolean isValid_usingIf(Change change) {
		if (change instanceof Created c) {
			return c.getInitialObject() == null && c.getNewObject() != null;
		}
		if (change instanceof Updated) {
			Updated u = (Updated) change;
			return u.getInitialObject() != null && u.getNewObject() != null;
		}
		if (change instanceof Deleted d) {
			return d.getInitialObject() != null && d.getNewObject() == null;
		}

		return false;
	}

	public static boolean isValid_usingSwitch(Change change) {
		
	}

}
