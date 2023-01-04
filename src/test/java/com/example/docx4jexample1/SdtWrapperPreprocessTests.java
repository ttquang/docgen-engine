package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.utils.XPathFactoryUtil;
import org.docx4j.wml.*;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.util.*;

@SpringBootTest
class SdtWrapperPreprocessTests {

    @Test
    void contextLoads() {
    }

    @Test
    void template_sdt() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document docData = db.parse(new File("data.xml"));
            docData.getDocumentElement().normalize();

            File doc = new File("staff-table.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            Queue<SdtElement> sdtElementQueue = new ArrayDeque<>();
            Stack<SdtWrapper> wrapperStack = new Stack<>();
            Queue<SdtWrapper> wrapperQueue = new ArrayDeque<>();
            List<Object> contents = List.copyOf(mainDocumentPart.getContent());

            for (Object content : contents) {
                proProcessContent(content, wrapperQueue, null);
            }

            for (SdtWrapper sdtWrapper : wrapperQueue) {
                wrapperStack.push(sdtWrapper);
            }

//            while (!sdtElementQueue.isEmpty()) {
//                SdtElement sdtElement = sdtElementQueue.poll();
//                wrapperStack.push(new SdtWrapper(sdtElement));
//            }
//
            while (!wrapperStack.empty()) {
                SdtWrapper sdtWrapper = wrapperStack.pop();
                System.out.println(sdtWrapper.getType());
                sdtWrapper.unWrap();
            }

            File exportFile = new File("staff1-result.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void proProcessContent(Object content, Queue<SdtWrapper> sdtWrapperQueue, SdtWrapper parentWrapper) {
        try {
            if (content instanceof JAXBElement) {
                content = ((JAXBElement) content).getValue();
            }

            if (!(content instanceof Text)) {
                if (content instanceof SdtBlock || content instanceof CTSdtRow) {
                    SdtElement sdtElement = (SdtElement) content;
                    SdtWrapper sdtWrapper = new SdtWrapper(sdtElement, parentWrapper);
                    sdtWrapperQueue.add(sdtWrapper);
                    for (Object child : sdtElement.getSdtContent().getContent()) {
                        proProcessContent(child, sdtWrapperQueue, sdtWrapper);
                    }
                } else {
                    ContentAccessor contentAccessor = (ContentAccessor) content;
                    for (Object child : contentAccessor.getContent()) {
                        proProcessContent(child, sdtWrapperQueue, parentWrapper);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
