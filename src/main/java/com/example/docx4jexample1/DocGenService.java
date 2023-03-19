package com.example.docx4jexample1;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.util.UUID;

@Service
public class DocGenService {
    @Autowired
    private TemplateContainer templateContainer;

    @Autowired
    private StoredProcedure1 storedProcedure1;

    public String execute(DocGenContext context) {
        try {
            String data = storedProcedure1.generateData(context);
            DataSource dataSource = DataSourceFactory.constructXMLDataSource(new ByteArrayInputStream(data.getBytes()));
            WordMLPackageWrapper wordMLPackageWrapper =
                    new WordMLPackageWrapper(templateContainer.get(context.getTemplate() + ".docx"), dataSource);
            wordMLPackageWrapper.process();
            wordMLPackageWrapper.save(new File(context.getTemplate() + UUID.randomUUID() + ".docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }

        return "";
    }

}
