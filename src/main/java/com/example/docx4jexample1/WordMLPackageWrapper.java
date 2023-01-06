package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.Stack;

public class WordMLPackageWrapper {
    private WordprocessingMLPackage wordMLPackage;
    private Stack<SdtElement> sdtBlockStack = new Stack<>();
    private Document data;

    public WordMLPackageWrapper(File templateFile, File dataFile) {
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            this.data = db.parse(dataFile);
            this.data.getDocumentElement().normalize();

            this.wordMLPackage = WordprocessingMLPackage.load(templateFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void process() {
        List<Object> contents = List.copyOf(wordMLPackage.getMainDocumentPart().getContent());
        for (Object content : contents) {
            processContent(content, "");
        }

        removeContentControl();
    }

    public void removeContentControl() {
        while (!sdtBlockStack.empty()) {
            SdtElement sdtElement = sdtBlockStack.pop();
            ContentAccessor contentAccessor = (ContentAccessor) sdtElement.getParent();
            int blockIndex = contentAccessor.getContent().indexOf(sdtElement);
            if (blockIndex >= 0) {
                contentAccessor.getContent().remove(blockIndex);
                contentAccessor.getContent().addAll(blockIndex, sdtElement.getSdtContent().getContent());
            }
        }
    }

    public void save(File file) {
        try {
            wordMLPackage.save(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void processContent(Object content, String parentDataPath) {
        try {
            XPath xPath = XPathFactory.newInstance().newXPath();
            if (isContentControlIncluding(content)) {
                if (content instanceof JAXBElement) {
                    content = ((JAXBElement) content).getValue();
                }

                if (content instanceof SdtElement) {
                    SdtElement sdtElement = (SdtElement) content;
                    sdtBlockStack.push(sdtElement);
                    String dataPatch = Utils.getTag(sdtElement);
                    String currentDataPath = parentDataPath + dataPatch;

                    if (Utils.isTextControl(sdtElement)) {
                        NodeList repeatCount = (NodeList) xPath.compile(currentDataPath).evaluate(data, XPathConstants.NODESET);
                        if (repeatCount.getLength() > 0) {
                            System.out.println(repeatCount.item(0).getTextContent());
                            sdtElement.getSdtPr().setShowingPlcHdr(false);
                            P p = Utils.createParagraphOfText(repeatCount.item(0).getTextContent());
                            sdtElement.getSdtContent().getContent().clear();
                            sdtElement.getSdtContent().getContent().add(p);

                        }
                    } else if (Utils.isLoopControl(sdtElement)) {
                        NodeList repeatCount = (NodeList) xPath.compile(currentDataPath).evaluate(data, XPathConstants.NODESET);
                        List<Object> templateContents = List.copyOf(sdtElement.getSdtContent().getContent());

                        for (int i = 1; i <= repeatCount.getLength(); i++) {
                            for (Object templateContent : templateContents) {
                                Object workingContent = XmlUtils.deepCopy(templateContent);
                                processContent(workingContent, currentDataPath + "[" + i + "]");
                                sdtElement.getSdtContent().getContent().add(workingContent);
                            }
                        }
                        sdtElement.getSdtContent().getContent().removeAll(templateContents);
                    } else if (Utils.isIfControl(sdtElement)) {
                        Map<String,Object> ifExpressionInput = XpathDataUtils.evaluate(data,dataPatch);
                        String ifExpression = Utils.getIfExpression(Utils.getAlias(sdtElement));
                        if (ConditionEvaluationUtils.evaluate(ifExpression,ifExpressionInput)) {
                            for (Object child : sdtElement.getSdtContent().getContent()) {
                                processContent(child, parentDataPath);
                            }
                        } else {
                            ContentAccessor contentAccessor = (ContentAccessor) sdtElement.getParent();
                            contentAccessor.getContent().remove(sdtElement);
                        }
                    }
                } else {
                    ContentAccessor contentAccessor = (ContentAccessor) content;

                    for (Object child : contentAccessor.getContent()) {
                        processContent(child, parentDataPath);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public boolean isContentControlIncluding(Object obj) {
        if (obj instanceof JAXBElement) {
            obj = ((JAXBElement) obj).getValue();
        }

        if (obj instanceof SdtBlock || obj instanceof CTSdtRow) {
            SdtElement sdtElement = (SdtElement) obj;
            if (Utils.isContentControl(sdtElement))
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

}
