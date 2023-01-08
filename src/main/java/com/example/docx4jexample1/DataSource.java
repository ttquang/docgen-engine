package com.example.docx4jexample1;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

public class DataSource {

    private Document data;
    private XPath xPath = XPathFactory.newInstance().newXPath();

    public DataSource(File file) {
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            this.data = db.parse(file);
            this.data.getDocumentElement().normalize();
        } catch (Exception e) {
            throw new IllegalArgumentException();
        }
    }

    public Map<String, Object> evaluateAsMap(String keyValuePairs) {
        Map<String, Object> result = new HashMap<>();
        if (!keyValuePairs.isBlank()) {
            try {
                String[] keyValuePairArr = keyValuePairs.split(";");
                for (String keyValuePair : keyValuePairArr) {
                    String s[] = keyValuePair.split("#");
                    String key = s[0];
                    String value = s[1];
                    NodeList nodeList = (NodeList) xPath.compile(value).evaluate(data, XPathConstants.NODESET);
                    if (nodeList.getLength() > 0) {
                        result.put(key, nodeList.item(0).getTextContent());
                    } else {
                        result.put(key, "");
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return result;
    }

    public String evaluateAsString(String xpath) {
        String result = "";
        if (!xpath.isBlank()) {
            try {
                NodeList nodeList = (NodeList) xPath.compile(xpath).evaluate(data, XPathConstants.NODESET);
                if (nodeList.getLength() > 0) {
                    result = nodeList.item(0).getTextContent();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return result;
    }

    public int count(String xpath) {
        int result = 0;
        if (!xpath.isBlank()) {
            try {
                NodeList nodeList = (NodeList) xPath.compile(xpath).evaluate(data, XPathConstants.NODESET);
                if (nodeList.getLength() > 0) {
                    result = nodeList.getLength();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return result;
    }
}
