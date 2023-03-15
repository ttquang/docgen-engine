package com.example.docx4jexample1;

import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.io.File;
import java.util.List;
import java.util.Optional;
import java.util.Stack;

public class WordMLPackageWrapper {
    private WordprocessingMLPackage wordMLPackage;
    private Stack<SdtElement> sdtBlockStack = new Stack<>();
    private DataSource dataSource;

    public WordMLPackageWrapper(WordprocessingMLPackage wordMLPackage, DataSource dataSource) {
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

    private void processContent(Object content, String parentDataPath) {
        try {
            if (isContentControlIncluding(content)) {
                content = Utils.getContent(content);

                if (content instanceof SdtElement) {
                    SdtElement sdtElement = (SdtElement) content;
                    sdtBlockStack.push(sdtElement);
                    String dataPatch = Utils.getTag(sdtElement);
                    String currentDataPath = parentDataPath + dataPatch;

                    if (Utils.isTextControl(sdtElement)) {
                        Optional<String> inputData = this.dataSource.getData(currentDataPath);
                        if (inputData.isPresent()) {
                            if (sdtElement instanceof SdtRun) {
                                SdtRun sdtRun = (SdtRun) sdtElement;
                                R run = (R) sdtRun.getSdtContent().getContent().get(0);
                                Text text = Utils.createText(inputData.get());
                                run.getContent().clear();
                                run.getContent().add(text);

                            } else {
                                sdtElement.getSdtPr().setShowingPlcHdr(false);
                                P p = Utils.createParagraphOfText(inputData.get());
                                sdtElement.getSdtContent().getContent().clear();
                                sdtElement.getSdtContent().getContent().add(p);
                            }
                        }
                    } else if (Utils.isLoopControl(sdtElement)) {
                        int repeatTimes = this.dataSource.count(currentDataPath);
                        List<Object> templateContents = List.copyOf(sdtElement.getSdtContent().getContent());

                        if (repeatTimes == 0) {
                            ContentAccessor contentAccessor = (ContentAccessor) sdtElement.getParent();
                            contentAccessor.getContent().remove(sdtElement);
                        } else {
                            for (int i = 1; i <= repeatTimes; i++) {
                                for (Object templateContent : templateContents) {
                                    Object workingContent = XmlUtils.deepCopy(templateContent);
                                    processContent(workingContent, currentDataPath + "[" + i + "]");
                                    sdtElement.getSdtContent().getContent().add(workingContent);
                                }
                            }
                        }

                        sdtElement.getSdtContent().getContent().removeAll(templateContents);
                    } else if (Utils.isIfControl(sdtElement)) {
                        String ifExpression = Utils.getIfExpression(Utils.getAlias(sdtElement));
                        if (this.dataSource.evaluateCondition(dataPatch,ifExpression)) {
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
        obj = Utils.getContent(obj);

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
