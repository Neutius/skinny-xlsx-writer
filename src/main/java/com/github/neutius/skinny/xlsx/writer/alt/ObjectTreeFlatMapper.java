package com.github.neutius.skinny.xlsx.writer.alt;

import com.github.neutius.skinny.xlsx.writer.interfaces.RowContentSupplier;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

public class ObjectTreeFlatMapper<T> implements RowContentSupplier {
    public static final int DEFAULT_DEPTH = 0;

    private final ArrayList<String> rowContent;
    private final T sourceObject;

    public ObjectTreeFlatMapper(T sourceObject) {
        this(sourceObject, DEFAULT_DEPTH);
    }

    public ObjectTreeFlatMapper(T sourceObject, int depth) {
        this.sourceObject = sourceObject;
        rowContent = new ArrayList<>();
        convertObjectToStringList(depth);
    }

    private void convertObjectToStringList(int depth) {
        List<Method> allGetMethods = findAllGetMethods();

        allGetMethods.sort(Comparator.comparing(Method::getName));

        allGetMethods.forEach(method -> rowContent.add(getStringValueFromGetMethod(method)));
    }

    private String getStringValueFromGetMethod(Method method) {
        String result = "";

        try {
            Object object = method.invoke(sourceObject);
            if (object instanceof String) {
                result = (String) object;
            }
            else {
                result = String.valueOf(object);
            }

        } catch (IllegalAccessException | InvocationTargetException e) {
            e.printStackTrace();
        }

        return result;
    }

    private List<Method> findAllGetMethods() {
        List<Method> methodList = Arrays.asList(sourceObject.getClass().getMethods());

        List<Method> allGetMethods = methodList.stream()
                .filter(this::isGetMethod)
                .filter(this::isNotGetClassMethod)
                .collect(Collectors.toList());

        System.out.println("methodList: " + methodList.stream().map(Method::getName).collect(Collectors.toList()));
        System.out.println("allGetMethods: " + allGetMethods.stream().map(Method::getName).collect(Collectors.toList()));
        return allGetMethods;
    }

    private boolean isNotGetClassMethod(Method method) {
        return !"getClass".equals(method.getName());
    }

    private boolean isGetMethod(Method method) {
        return method.getParameterCount() == DEFAULT_DEPTH && (method.getName().startsWith("get") || method.getName().startsWith("is"));
    }

    @Override
    public List<String> get() {
        return rowContent;
    }
}
