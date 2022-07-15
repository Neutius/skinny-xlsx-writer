package com.github.neutius.skinny.xlsx.test;

public class StringValueTestObjectWithPrivateGetter extends StringValueTestObject {

    private final String combined;

    public StringValueTestObjectWithPrivateGetter(String name, String id, String favouriteColor) {
        super(name, id, favouriteColor);
        combined = getCombined();
    }

    private String getCombined() {
        return getName() + getId() + getFavouriteColor();
    }

    @Override
    public String toString() {
        return "StringValueTestObjectWithPrivateGetter{" +
                "combined='" + combined + '\'' +
                '}';
    }
}
