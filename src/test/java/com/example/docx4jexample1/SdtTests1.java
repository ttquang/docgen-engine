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

@SpringBootTest
class SdtTests1 {

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

            File doc = new File("staff1.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

            for (Object content : mainDocumentPart.getContent()) {
                processContent(content, docData, "");
            }

//            File exportFile = new File("staff1-result.docx");
//            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void processContent(Object content, Document docData, String parentDataPath) {
        try {
            if (content instanceof SdtBlock) {
                SdtBlock block = (SdtBlock) content;
                ContentAccessor blockParent = (ContentAccessor) block.getParent();
                int blockIndex = blockParent.getContent().indexOf(block);
                SdtBlockWrapper blockWrapper = new SdtBlockWrapper(block,parentDataPath);
                blockParent.getContent().remove(block);
                blockParent.getContent().add(blockIndex, blockWrapper);


                SdtElement sdtElement = (SdtElement) content;
                String dataPatch = getTag(sdtElement);
                String currentDataPath = parentDataPath + dataPatch;

                if (isTextControl(sdtElement)) {

                } else if (isLoopControl(sdtElement)) {

                }
            } else {
                ContentAccessor contentAccessor = (ContentAccessor) content;
                for (Object child : contentAccessor.getContent()) {
                    processContent(child, docData, parentDataPath);
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

        if (obj instanceof SdtBlock) {
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
                    if ("Loop".equals(alias.getVal()) || "Text".equals(alias.getVal()))
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

    private String getTag(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof Tag) {
                return ((Tag) prObj).getVal();
            }
        }
        return "";
    }

}
