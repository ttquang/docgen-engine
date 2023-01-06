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
            WordMLPackageWrapper wordMLPackageWrapper = new WordMLPackageWrapper(new File("staff-table.docx"),new File("data.xml"));
            wordMLPackageWrapper.process();
            wordMLPackageWrapper.save(new File("staff1-result.docx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
