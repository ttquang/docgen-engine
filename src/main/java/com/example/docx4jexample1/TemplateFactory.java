package com.example.docx4jexample1;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

public class TemplateFactory {
    static private Map<String, WordprocessingMLPackage> factory = new HashMap<>();

    static void load() {
        try {
            File templateDir = new ClassPathResource("doc-gen-template").getFile();
            for (File templateFile : templateDir.listFiles()) {
                factory.put(templateFile.getName(), WordprocessingMLPackage.load(templateFile));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    static WordprocessingMLPackage get(String templateName) {
        return (WordprocessingMLPackage) factory.get(templateName).clone();
    }
}
