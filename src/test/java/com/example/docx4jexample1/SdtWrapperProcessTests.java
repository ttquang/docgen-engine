package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.utils.XPathFactoryUtil;
import org.docx4j.wml.*;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.io.File;
import java.util.*;

@SpringBootTest
class SdtWrapperProcessTests {

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

            for (Object content : mainDocumentPart.getContent()) {
                processContent(content, docData, "");
            }

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

    private void processContent(Object content, Document docData, String parentDataPath) {
        try {
            XPath xPath = XPathFactory.newInstance().newXPath();
            if (isContentControlIncluding(content)) {
                if (content instanceof JAXBElement) {
                    content = ((JAXBElement) content).getValue();
                }

                if (content instanceof SdtWrapper) {
                    SdtWrapper sdtWrapper = (SdtWrapper) content;
                    SdtElement sdtElement = sdtWrapper.getSdt();
                    String dataPatch = sdtWrapper.getDatasource();
                    String currentDataPath = parentDataPath + dataPatch;

                    if (sdtWrapper.getType() == SdtType.TEXT) {
                        NodeList repeatCount = (NodeList) xPath.compile(currentDataPath).evaluate(docData, XPathConstants.NODESET);
                        if (repeatCount.getLength() > 0) {
                            System.out.println(repeatCount.item(0).getTextContent());
                            sdtElement.getSdtPr().setShowingPlcHdr(false);
                            P p = Utils.createParagraphOfText(repeatCount.item(0).getTextContent());
                            sdtElement.getSdtContent().getContent().clear();
                            sdtElement.getSdtContent().getContent().add(p);

                        }
                    } else if (sdtWrapper.getType() == SdtType.LOOP) {
                        NodeList repeatCount = (NodeList) xPath.compile(currentDataPath).evaluate(docData, XPathConstants.NODESET);
                        List<Object> templateContents = List.copyOf(sdtElement.getSdtContent().getContent());

                        for (int i = 1; i <= repeatCount.getLength(); i++) {
                            for (Object templateContent : templateContents) {
                                Object workingContent = XmlUtils.deepCopy(templateContent);
                                processContent(workingContent, docData, currentDataPath + "[" + i + "]");
                                sdtElement.getSdtContent().getContent().add(workingContent);
                            }
                        }
                        sdtElement.getSdtContent().getContent().removeAll(templateContents);
                    } else if (sdtWrapper.getType() == SdtType.IF) {
                        for (Object child : sdtElement.getSdtContent().getContent()) {
                            processContent(child, docData, parentDataPath);
                        }
//                        ContentAccessor contentAccessor = (ContentAccessor) sdtElement.getParent();
//                        contentAccessor.getContent().remove(sdtElement);
                    }
                } else {
                    ContentAccessor contentAccessor = (ContentAccessor) content;
                    for (Object child : contentAccessor.getContent()) {
                        processContent(child, docData, parentDataPath);
                    }
                }
            }
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

    private boolean isContentControlIncluding(Object obj) {
        if (obj instanceof JAXBElement) {
            obj = ((JAXBElement) obj).getValue();
        }

        if (obj instanceof SdtWrapper) {
            SdtWrapper sdtWrapper = (SdtWrapper) obj;
            if (Objects.nonNull(sdtWrapper.getType()))
                return true;
        } else if (obj instanceof ContentAccessor) {
            List<Object> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                if (isContentControlIncluding(child))
                    return true;
            }
        }
        return false;
    }
}
