package com.example.docx4jexample1;

import org.docx4j.utils.XPathFactoryUtil;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;

@SpringBootTest
class WordMLPackageWrapperTest {

    @Test
    void contextLoads() {
    }

    @Test
    void template_sdt() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
        try {
            TemplateFactory.load();
            File dataFile = new File("CC_Application_Form-data.xml");
            DataSource dataSource = DataSourceFactory.constructXMLDataSource(dataFile);
            WordMLPackageWrapper wordMLPackageWrapper =
//                    new WordMLPackageWrapper(new File("Cheque_Issuance_Form.docx"),new File("Cheque_Issuance_Form-data.xml"));
                    new WordMLPackageWrapper(TemplateFactory.get("CC_Application_Form_new.docx"), dataSource);
            wordMLPackageWrapper.process();
            wordMLPackageWrapper.save(new File("CC_Application_Form-result.docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template_sdt_1() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
        try {
            TemplateFactory.load();
            File dataFile = new File("CC_Application_Form-data.xml");
            DataSource dataSource = DataSourceFactory.constructXMLDataSource(dataFile);
            WordMLPackageWrapper wordMLPackageWrapper =
//                    new WordMLPackageWrapper(new File("Cheque_Issuance_Form.docx"),new File("Cheque_Issuance_Form-data.xml"));
                    new WordMLPackageWrapper(TemplateFactory.get("CC_Fixing_Document_Cover_Page.docx"), dataSource);
            wordMLPackageWrapper.process();
            wordMLPackageWrapper.save(new File("CC_Fixing_Document_Cover_Page-result.docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
