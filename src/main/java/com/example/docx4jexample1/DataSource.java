package com.example.docx4jexample1;

import java.util.List;
import java.util.Optional;

public interface DataSource {
    Optional<String> getData(String name);
    List<String> getData2(String name);
    int count(String name);
    boolean evaluateCondition(String dataPatch, String conditionExpression);
}
