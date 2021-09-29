package com.github.neutius.skinny.xlsx.writer.alt;

public class StringValueTestObjectWithValueObjectFields extends StringValueTestObject {

    private final StringValueTestObject partner;

    public StringValueTestObjectWithValueObjectFields(String name, String id, String favouriteColor, StringValueTestObject partner) {
        super(name, id, favouriteColor);
        this.partner = partner;
    }

    public StringValueTestObject getPartner() {
        return partner;
    }
}
