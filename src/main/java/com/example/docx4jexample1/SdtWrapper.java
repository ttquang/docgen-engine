package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.wml.*;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class SdtWrapper {
    private SdtElement sdt;
    private SdtType type;
    private String datasource = "";
    private boolean contentControl = false;
    private SdtWrapper parentWrapper;
    private List<SdtWrapper> childWrappers = new ArrayList<>();
    private ContentAccessor parent;

    public SdtWrapper(SdtElement sdt, SdtWrapper parentWrapper) {
        this.sdt = sdt;
        this.parentWrapper = parentWrapper;
        for (Object prObj : sdt.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    this.type = SdtType.valueOfLabel(alias.getVal());
                }
            } else if (prObj instanceof Tag) {
                this.datasource += ((Tag) prObj).getVal();
            }
        }
        if (Objects.nonNull(this.parentWrapper)) {
            this.parentWrapper.childWrappers.add(this);
        }
        this.parent = (ContentAccessor) sdt.getParent();
        int sdtIndex = parent.getContent().indexOf(sdt);
        parent.getContent().remove(sdt);
        parent.getContent().add(sdtIndex, this);
        if (Objects.nonNull(this.type)) {
            this.contentControl = true;
            SdtWrapper tmpParentWrapper = this.parentWrapper;
            while (Objects.nonNull(tmpParentWrapper)) {
                tmpParentWrapper.contentControl = true;
                tmpParentWrapper = tmpParentWrapper.parentWrapper;
            }
        }
    }

    public SdtElement getSdt() {
        return sdt;
    }

    public SdtType getType() {
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

    public SdtWrapper getParentWrapper() {
        return parentWrapper;
    }

    public void setParentWrapper(SdtWrapper parentWrapper) {
        this.parentWrapper = parentWrapper;
    }

    public List<SdtWrapper> getChildWrappers() {
        return childWrappers;
    }

    public void unWrap() {
        int blockIndex = parent.getContent().indexOf(this);
        if (blockIndex >= 0) {
            parent.getContent().remove(blockIndex);
            parent.getContent().addAll(blockIndex, sdt.getSdtContent().getContent());
        }
    }
}
