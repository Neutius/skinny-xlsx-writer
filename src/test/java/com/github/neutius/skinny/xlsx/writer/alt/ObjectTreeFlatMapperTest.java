package com.github.neutius.skinny.xlsx.writer.alt;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import static org.assertj.core.api.Assertions.assertThat;

class ObjectTreeFlatMapperTest {

    private static final String NAME = "name";
    private static final String ID = "ID";
    private static final String COLOR = "color";
    private static final String SONG = "song";
    private static final String MOVIE = "movie";

    @Test
    void onlyStringFields_returnsStringList() {
        StringValueTestObject input = new StringValueTestObject(NAME, ID, COLOR);

        ObjectTreeFlatMapper<StringValueTestObject> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).isNotNull().isNotEmpty().hasSize(3);
    }

    @Test
    void onlyStringFields_subClass_returnsStringList() {
        StringValueTestObjectSubClass input = new StringValueTestObjectSubClass(NAME, ID, COLOR, SONG, MOVIE);

        ObjectTreeFlatMapper<StringValueTestObjectSubClass> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).isNotNull().isNotEmpty().hasSize(5);
    }

    @Test
    void onlyStringFields_hasPrivateGetMethod_returnsStringList() {
        StringValueTestObjectWithPrivateGetter input = new StringValueTestObjectWithPrivateGetter(NAME, ID, COLOR);

        ObjectTreeFlatMapper<StringValueTestObjectWithPrivateGetter> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).isNotNull().isNotEmpty().hasSize(3);
    }

    @Test
    void onlyStringFields_returnsValuesOfFields() {
        StringValueTestObject input = new StringValueTestObject(NAME, ID, COLOR);

        ObjectTreeFlatMapper<StringValueTestObject> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).containsExactlyInAnyOrder(NAME, ID, COLOR);
    }

    @Test
    void onlyStringFields_subClass_returnsValuesOfFields() {
        StringValueTestObjectSubClass input = new StringValueTestObjectSubClass(NAME, ID, COLOR, SONG, MOVIE);

        ObjectTreeFlatMapper<StringValueTestObjectSubClass> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).containsExactlyInAnyOrder(NAME, ID, COLOR, SONG, MOVIE);
    }

    @Test
    void onlyStringFields_hasPrivateGetMethod_returnsValuesOfFields() {
        StringValueTestObjectWithPrivateGetter input = new StringValueTestObjectWithPrivateGetter(NAME, ID, COLOR);

        ObjectTreeFlatMapper<StringValueTestObjectWithPrivateGetter> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).containsExactlyInAnyOrder(NAME, ID, COLOR);
    }

    @Test
    void onlyStringFields_subClass_valuesAreSortedByFieldName() {
        StringValueTestObjectSubClass input = new StringValueTestObjectSubClass(NAME, ID, COLOR, SONG, MOVIE);

        ObjectTreeFlatMapper<StringValueTestObjectSubClass> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).containsExactly(COLOR, MOVIE, SONG, ID, NAME);
    }

    @Test
    void stringAndPrimitiveFields_primitivesAreConvertedToString() {
        StringsAndPrimitivesTestObject input = new StringsAndPrimitivesTestObject(NAME, ID, COLOR, 23, 1.83, true, false);

        ObjectTreeFlatMapper<StringsAndPrimitivesTestObject> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).isNotNull().isNotEmpty().hasSize(7);
    }

    @Test
    void stringAndPrimitiveFields_primitiveValuesAreRepresentedAsString() {
        StringsAndPrimitivesTestObject input = new StringsAndPrimitivesTestObject(NAME, ID, COLOR, 23, 1.83, true, false);

        ObjectTreeFlatMapper<StringsAndPrimitivesTestObject> testSubject = new ObjectTreeFlatMapper<>(input);

        assertThat(testSubject.getRowContent()).containsExactlyInAnyOrder(NAME, ID, COLOR, "23", "1.83", "true", "false");
    }

    @Disabled("Failing test for experimental code")
    @Test
    void objectWithObjectField_depthZero_fieldsOfNestedObjectAreConvertedToString() {
        StringValueTestObject partner = new StringValueTestObject("Name-2", "ID-2", "Color-2");
        StringValueTestObjectWithValueObjectFields input = new StringValueTestObjectWithValueObjectFields(NAME, ID, COLOR, partner);

        ObjectTreeFlatMapper<StringValueTestObjectWithValueObjectFields> testSubject = new ObjectTreeFlatMapper<>(input, 0);

        assertThat(testSubject.getRowContent()).isNotNull().isNotEmpty().hasSize(6);
    }

}
