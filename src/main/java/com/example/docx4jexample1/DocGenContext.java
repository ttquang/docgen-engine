package com.example.docx4jexample1;

import jakarta.xml.bind.annotation.XmlRootElement;

import java.io.Serializable;

@XmlRootElement
public class DocGenContext implements Serializable {

    private String template;
    private String oracleXML;
    private String applicationNo;

    public DocGenContext() {}

    public String getTemplate() {
        return template;
    }

    public void setTemplate(String template) {
        this.template = template;
    }

    public String getOracleXML() {
        return oracleXML;
    }

    public void setOracleXML(String oracleXML) {
        this.oracleXML = oracleXML;
    }

    public String getApplicationNo() {
        return applicationNo;
    }

    public void setApplicationNo(String applicationNo) {
        this.applicationNo = applicationNo;
    }
}
