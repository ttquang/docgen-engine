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
import java.util.List;
import java.util.Stack;

@SpringBootTest
class SdtTests {

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
            Stack<SdtElement> sdtBlockStack = new Stack<>();

            for (Object content : mainDocumentPart.getContent()) {
                processContent(content, docData, "", sdtBlockStack);
            }

            while (!sdtBlockStack.empty()) {
                SdtElement sdtElement = sdtBlockStack.pop();
                ContentAccessor contentAccessor = (ContentAccessor) sdtElement.getParent();
                int blockIndex = contentAccessor.getContent().indexOf(sdtElement);
                if (blockIndex >= 0) {
                    contentAccessor.getContent().remove(blockIndex);
                    contentAccessor.getContent().addAll(blockIndex, sdtElement.getSdtContent().getContent());
                }
            }

            File exportFile = new File("staff1-result.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void processContent(Object content, Document docData, String parentDataPath, Stack<SdtElement> sdtBlockStack) {
        try {
            XPath xPath = XPathFactory.newInstance().newXPath();
            if (isContentControlIncluding(content)) {
                if (content instanceof JAXBElement) {
                    content = ((JAXBElement) content).getValue();
                }

                if (content instanceof SdtBlock || content instanceof CTSdtRow) {
                    SdtElement sdtElement = (SdtElement) content;
                    sdtBlockStack.push(sdtElement);
                    String dataPatch = getTag(sdtElement);
                    String currentDataPath = parentDataPath + dataPatch;

                    if (isTextControl(sdtElement)) {
                        NodeList repeatCount = (NodeList) xPath.compile(currentDataPath).evaluate(docData, XPathConstants.NODESET);
                        if (repeatCount.getLength() > 0) {
                            System.out.println(repeatCount.item(0).getTextContent());
                            sdtElement.getSdtPr().setShowingPlcHdr(false);
                            P p = Utils.createParagraphOfText(repeatCount.item(0).getTextContent());
                            sdtElement.getSdtContent().getContent().clear();
                            sdtElement.getSdtContent().getContent().add(p);

                        }
                    } else if (isLoopControl(sdtElement)) {
                        NodeList repeatCount = (NodeList) xPath.compile(currentDataPath).evaluate(docData, XPathConstants.NODESET);
                        List<Object> templateContents = List.copyOf(sdtElement.getSdtContent().getContent());

                        for (int i = 1; i <= repeatCount.getLength(); i++) {
                            for (Object templateContent : templateContents) {
                                Object workingContent = XmlUtils.deepCopy(templateContent);
                                processContent(workingContent, docData, currentDataPath + "[" + i + "]", sdtBlockStack);
                                sdtElement.getSdtContent().getContent().add(workingContent);
                            }
                        }
                        sdtElement.getSdtContent().getContent().removeAll(templateContents);
                    } else if (isIfControl(sdtElement)) {
                        for (Object child:sdtElement.getSdtContent().getContent()) {
                            processContent(child, docData, parentDataPath, sdtBlockStack);
                        }
//                        ContentAccessor contentAccessor = (ContentAccessor) sdtElement.getParent();
//                        contentAccessor.getContent().remove(sdtElement);
                    }
                } else {
                    ContentAccessor contentAccessor;
                    if (content instanceof JAXBElement) {
                        contentAccessor = (ContentAccessor) ((JAXBElement) content).getValue();
                    } else {
                        contentAccessor = (ContentAccessor) content;
                    }

                    for (Object child : contentAccessor.getContent()) {
                        processContent(child, docData, parentDataPath, sdtBlockStack);
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

        if (obj instanceof SdtBlock || obj instanceof CTSdtRow) {
            SdtElement sdtElement = (SdtElement) obj;
            if (isContentControl(sdtElement))
                return true;
        } else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                if (isContentControlIncluding(child))
                    return true;
            }
        }
        return false;
    }

    private boolean isContentControl(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    if ("Loop".equals(alias.getVal()) || "Text".equals(alias.getVal()) || "If".equals(alias.getVal()))
                        return true;
                }
            }
        }
        return false;
    }

    private boolean isTextControl(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    if ("Text".equals(alias.getVal()))
                        return true;
                }
            }
        }
        return false;
    }

    private boolean isLoopControl(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    if ("Loop".equals(alias.getVal()))
                        return true;
                }
            }
        }
        return false;
    }

    private boolean isIfControl(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    if ("If".equals(alias.getVal()))
                        return true;
                }
            }
        }
        return false;
    }

    private String getTag(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof Tag) {
                return ((Tag) prObj).getVal();
            }
        }
        return "";
    }

}
