package com.example.docx4jexample1;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

@Configuration
public class DocGenConfiguration {

    @Bean
    public TemplateContainer templateContainer() {
        Map<String, WordprocessingMLPackage> templates = new HashMap<>();
        try {
            File templateDir = new ClassPathResource("doc-gen-template").getFile();
            for (File templateFile : templateDir.listFiles()) {
                templates.put(templateFile.getName(), WordprocessingMLPackage.load(templateFile));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return new TemplateContainer(templates);
    }

}
