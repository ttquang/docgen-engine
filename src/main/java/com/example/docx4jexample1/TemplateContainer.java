package com.example.docx4jexample1;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.util.Map;

public class TemplateContainer {

    private Map<String, WordprocessingMLPackage> templates;

    public TemplateContainer(Map<String, WordprocessingMLPackage> templates) {
        this.templates = templates;
    }

    public WordprocessingMLPackage get(String templateName) {
        return (WordprocessingMLPackage) templates.get(templateName).clone();
    }

}
