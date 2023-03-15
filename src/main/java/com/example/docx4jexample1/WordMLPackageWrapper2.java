package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.Stack;

public class WordMLPackageWrapper2 {
    private WordprocessingMLPackage wordMLPackage;
    private Stack<SdtElement> sdtBlockStack = new Stack<>();
    private Document data;
    private DataSource dataSource;

    public WordMLPackageWrapper2(File templateFile, File dataFile) {
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
//            this.data = db.parse(dataFile);
//            this.data.getDocumentElement().normalize();

            WordprocessingMLPackage templatePackage = WordprocessingMLPackage.load(templateFile);
            this.wordMLPackage = (WordprocessingMLPackage) templatePackage.clone();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public WordMLPackageWrapper2(WordprocessingMLPackage wordMLPackage, DataSource dataSource) {
        this.dataSource = dataSource;
        this.wordMLPackage = wordMLPackage;
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

    public void savePDF(File file) {
        try {
            FileOutputStream os = new FileOutputStream(file);
            Docx4J.toPDF(wordMLPackage,os);
            os.flush();
            os.close();
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
                            if (sdtElement instanceof SdtRun) {
                                SdtRun sdtRun = (SdtRun) sdtElement;
                                R run = (R) sdtRun.getSdtContent().getContent().get(0);
                                Text text = Utils.createText(repeatCount.item(0).getTextContent());
                                run.getContent().clear();
                                run.getContent().add(text);

                            } else {
                                sdtElement.getSdtPr().setShowingPlcHdr(false);
                                P p = Utils.createParagraphOfText(repeatCount.item(0).getTextContent());
                                sdtElement.getSdtContent().getContent().clear();
                                sdtElement.getSdtContent().getContent().add(p);
                            }
                        }
                    } else if (Utils.isLoopControl(sdtElement)) {
                        NodeList repeatCount = (NodeList) xPath.compile(currentDataPath).evaluate(data, XPathConstants.NODESET);
                        List<Object> templateContents = List.copyOf(sdtElement.getSdtContent().getContent());

                        if (repeatCount.getLength() == 0) {
                            ContentAccessor contentAccessor = (ContentAccessor) sdtElement.getParent();
                            contentAccessor.getContent().remove(sdtElement);
                        } else {
                            for (int i = 1; i <= repeatCount.getLength(); i++) {
                                for (Object templateContent : templateContents) {
                                    Object workingContent = XmlUtils.deepCopy(templateContent);
                                    processContent(workingContent, currentDataPath + "[" + i + "]");
                                    sdtElement.getSdtContent().getContent().add(workingContent);
                                }
                            }
                        }

                        sdtElement.getSdtContent().getContent().removeAll(templateContents);
                    } else if (Utils.isIfControl(sdtElement)) {
//                        Map<String,Object> ifExpressionInput = XpathDataUtils.evaluate(data,dataPatch);
//                        String ifExpression = Utils.getIfExpression(Utils.getAlias(sdtElement));
//                        if (ConditionEvaluationUtils.evaluate(ifExpression,ifExpressionInput)) {
//                            for (Object child : sdtElement.getSdtContent().getContent()) {
//                                processContent(child, parentDataPath);
//                            }
//                        } else {
//                            ContentAccessor contentAccessor = (ContentAccessor) sdtElement.getParent();
//                            contentAccessor.getContent().remove(sdtElement);
//                        }
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

        if (obj instanceof SdtElement) {
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
