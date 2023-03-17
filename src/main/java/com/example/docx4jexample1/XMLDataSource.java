package com.example.docx4jexample1;

import org.w3c.dom.Document;
import org.w3c.dom.NodeList;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class XMLDataSource implements DataSource {

    private Document document;
    private XPath xPath = XPathFactory.newInstance().newXPath();

    public XMLDataSource(Document document) {
        this.document = document;
    }

    @Override
    public Optional<String> getData(String name) {
        Optional<String> result = Optional.empty();
        try {
            NodeList repeatCount = (NodeList) this.xPath.compile(name).evaluate(document, XPathConstants.NODESET);
            if (repeatCount.getLength() > 0) {
                result = Optional.of(repeatCount.item(0).getTextContent());
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return result;
        }
    }

    @Override
    public List<String> getData2(String name) {
        return null;
    }

    @Override
    public int count(String name) {
        int result = 0;
        try {
            NodeList repeatCount = (NodeList) this.xPath.compile(name).evaluate(document, XPathConstants.NODESET);
            if (repeatCount.getLength() > 0) {
                result = repeatCount.getLength();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return result;
        }
    }

    @Override
    public boolean evaluateCondition(String dataPatch, String ifExpression) {
        boolean result = false;
        try {
            Map<String, Object> ifExpressionInput = new HashMap<>();
            if (!dataPatch.isBlank()) {
                String[] keyValuePairArr = dataPatch.split(";");
                for (String keyValuePair : keyValuePairArr) {
                    String s[] = keyValuePair.split("#");
                    String key = s[0];
                    String value = s[1];
                    NodeList nodeList = (NodeList) this.xPath.compile(value).evaluate(document, XPathConstants.NODESET);
                    if (nodeList.getLength() > 0) {
                        ifExpressionInput.put(key, nodeList.item(0).getTextContent());
                    } else {
                        ifExpressionInput.put(key, "");
                    }
                }

                result = ConditionEvaluationUtils.evaluate(ifExpression, ifExpressionInput);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return result;
        }
    }

    @Override
    public boolean evaluateCondition(String parentDataPatch, String dataPatch, String conditionExpression) {
        boolean result = false;
        try {
            Map<String, Object> ifExpressionInput = new HashMap<>();
            if (!dataPatch.isBlank()) {
                String[] keyValuePairArr = dataPatch.split(";");
                for (String keyValuePair : keyValuePairArr) {
                    String s[] = keyValuePair.split("#");
                    String key = s[0];
                    if (key.startsWith("./")) {
                        key = key.replace(".", parentDataPatch);
                    }
                    String value = s[1];
                    NodeList nodeList = (NodeList) this.xPath.compile(value).evaluate(document, XPathConstants.NODESET);
                    if (nodeList.getLength() > 0) {
                        ifExpressionInput.put(key, nodeList.item(0).getTextContent());
                    } else {
                        ifExpressionInput.put(key, "");
                    }
                }

                result = ConditionEvaluationUtils.evaluate(conditionExpression, ifExpressionInput);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return result;
        }
    }

}
