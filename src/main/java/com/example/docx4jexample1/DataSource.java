package com.example.docx4jexample1;

import java.util.List;
import java.util.Optional;

public interface DataSource {
    Optional<String> getData(String name);
    Optional<String> getData(String parentDataPatch, String dataPatch);
    List<String> getData2(String name);
    int count(String name);
    boolean evaluateSimpleIfCondition(String parentDataPatch, String dataPatch);
    boolean evaluateSimpleIfNotCondition(String parentDataPatch, String dataPatch);
}
