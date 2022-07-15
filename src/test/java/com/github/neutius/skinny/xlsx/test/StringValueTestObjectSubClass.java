package com.github.neutius.skinny.xlsx.test;

public class StringValueTestObjectSubClass extends StringValueTestObject {

    private final String favouriteSong;
    private final String favouriteMovie;

    public StringValueTestObjectSubClass(String name, String id, String favouriteColor, String favouriteSong, String favouriteMovie) {
        super(name, id, favouriteColor);
        this.favouriteSong = favouriteSong;
        this.favouriteMovie = favouriteMovie;
    }

    public String getFavouriteSong() {
        return favouriteSong;
    }

    public String getFavouriteMovie() {
        return favouriteMovie;
    }
}
