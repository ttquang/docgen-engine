package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.SdtBlock;
import org.docx4j.wml.SdtPr;
import org.docx4j.wml.Tag;

public class SdtBlockWrapper {
    private SdtBlock block;
    private SdtBlockType type;
    private String datasource;
    private boolean contentControl = false;
    private SdtBlockWrapper parentWrapper;

    public SdtBlockWrapper(SdtBlock block, String parentDatasource) {
        this.block = block;
        this.datasource = parentDatasource;
        for (Object prObj : block.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    this.type = SdtBlockType.valueOfLabel(alias.getVal());
                }
            } else if (prObj instanceof Tag) {
                this.datasource += ((Tag) prObj).getVal();
            }

            ContentAccessor blockParent = (ContentAccessor) block.getParent();
        }
    }

    public SdtBlock getBlock() {
        return block;
    }

    public SdtBlockType getType() {
        return type;
    }

    public String getDatasource() {
        return datasource;
    }

    public boolean isContentControl() {
        return contentControl;
    }

    public void setContentControl(boolean contentControl) {
        this.contentControl = contentControl;
    }

    public SdtBlockWrapper getParentWrapper() {
        return parentWrapper;
    }

    public void setParentWrapper(SdtBlockWrapper parentWrapper) {
        this.parentWrapper = parentWrapper;
    }
}
