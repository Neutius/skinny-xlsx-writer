package com.github.neutius.skinny.xlsx.writer.alt;

public class StringsAndPrimitivesTestObject extends StringValueTestObject {

    private final int age;
    private final double height;
    private final boolean catPerson;
    private final boolean dogPerson;

    public StringsAndPrimitivesTestObject(String name, String id, String favouriteColor,
                                          int age, double height, boolean catPerson, boolean dogPerson) {
        super(name, id, favouriteColor);
        this.age = age;
        this.height = height;
        this.catPerson = catPerson;
        this.dogPerson = dogPerson;
    }

    public int getAge() {
        return age;
    }

    public double getHeight() {
        return height;
    }

    public boolean isCatPerson() {
        return catPerson;
    }

    public boolean isDogPerson() {
        return dogPerson;
    }
}
