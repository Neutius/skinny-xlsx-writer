package com.github.neutius.skinny.xlsx.writer.alt;

public class StringValueTestObject {

    private final String name;
    private final String id;
    private final String favouriteColor;

    public StringValueTestObject(String name, String id, String favouriteColor) {
        this.name = name;
        this.id = id;
        this.favouriteColor = favouriteColor;
    }

    public String getName() {
        return name;
    }

    public String getId() {
        return id;
    }

    public String getFavouriteColor() {
        return favouriteColor;
    }
}
