package com.example.docx4jexample1;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.UUID;

@SpringBootTest
class WordMLPackageWrapperTest {

    @Autowired
    private TemplateContainer templateContainer;

    @Test
    void contextLoads() {
    }

    @Test
    void template_sdt() {
        try {
            File dataFile = new File("CC_Application_Form-data.xml");
            DataSource dataSource = DataSourceFactory.constructXMLDataSource(dataFile);
            WordMLPackageWrapper wordMLPackageWrapper =
                    new WordMLPackageWrapper(templateContainer.get("CC_Application_Form_new.docx"), dataSource);
            wordMLPackageWrapper.process();
            wordMLPackageWrapper.save(new File("CC_Application_Form-" + UUID.randomUUID() + ".docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template_sdt_1() {
        try {
            File dataFile = new File("CC_Application_Form-data.xml");
            DataSource dataSource = DataSourceFactory.constructXMLDataSource(dataFile);
            WordMLPackageWrapper wordMLPackageWrapper =
                    new WordMLPackageWrapper(templateContainer.get("CC_Fixing_Document_Cover_Page.docx"), dataSource);
            wordMLPackageWrapper.process();
            wordMLPackageWrapper.save(new File("CC_Fixing_Document_Cover_Page-result.docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
