package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.Stack;

public class WordMLPackageWrapper {
    private WordprocessingMLPackage wordMLPackage;
    private Stack<SdtElement> sdtElementStack = new Stack<>();
    private DataSource ds;

    public WordMLPackageWrapper(File templateFile, File dataFile) {
        try {
            this.ds = new DataSource(dataFile);
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
        while (!sdtElementStack.empty()) {
            SdtElement sdtElement = sdtElementStack.pop();
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
            if (isContentControlIncluding(content)) {
                if (content instanceof JAXBElement) {
                    content = ((JAXBElement) content).getValue();
                }

                if (content instanceof SdtElement) {
                    SdtElement sdtElement = (SdtElement) content;
                    sdtElementStack.push(sdtElement);
                    String dataPath = sdtElement.getSdtPr().getTag().getVal();
                    String currentDataPath = parentDataPath + dataPath;

                    if (Utils.isTextControl(sdtElement)) {
                        sdtElement.getSdtPr().getTag().setVal(currentDataPath);
                        handleText(sdtElement);
                    } else if (Utils.isLoopControl(sdtElement)) {
                        sdtElement.getSdtPr().getTag().setVal(currentDataPath);
                        handleLoop(sdtElement);
                    } else if (Utils.isIfControl(sdtElement)) {
                        handleIf(sdtElement, parentDataPath);
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

    public void handleText(SdtElement contentControl) {
        String value = ds.evaluateAsString(contentControl.getSdtPr().getTag().getVal());
        contentControl.getSdtPr().setShowingPlcHdr(false);
        P p = Utils.createParagraphOfText(value);
        contentControl.getSdtContent().getContent().clear();
        contentControl.getSdtContent().getContent().add(p);
    }

    public void handleLoop(SdtElement contentControl) {
        int count = ds.count(contentControl.getSdtPr().getTag().getVal());
        List<Object> templateContents = List.copyOf(contentControl.getSdtContent().getContent());

        for (int i = 1; i <= count; i++) {
            for (Object templateContent : templateContents) {
                Object workingContent = XmlUtils.deepCopy(templateContent);
                processContent(workingContent, contentControl.getSdtPr().getTag().getVal() + "[" + i + "]");
                contentControl.getSdtContent().getContent().add(workingContent);
            }
        }
        contentControl.getSdtContent().getContent().removeAll(templateContents);
    }

    public void handleIf(SdtElement contentControl, String parentDataPath) {
        Map<String, Object> ifExpressionInput = ds.evaluateAsMap(contentControl.getSdtPr().getTag().getVal());
        String ifExpression = Utils.getIfExpression(Utils.getAlias(contentControl));
        if (ConditionEvaluationUtils.evaluate(ifExpression, ifExpressionInput)) {
            for (Object child : contentControl.getSdtContent().getContent()) {
                processContent(child, parentDataPath);
            }
        } else {
            ContentAccessor contentAccessor = (ContentAccessor) contentControl.getParent();
            contentAccessor.getContent().remove(contentControl);
        }
    }

}
