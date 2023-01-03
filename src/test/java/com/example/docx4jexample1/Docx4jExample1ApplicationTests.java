package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.model.datastorage.migration.VariablePrepare;
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
import java.io.FileInputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@SpringBootTest
class Docx4jExample1ApplicationTests {

    @Test
    void contextLoads() {
    }

    @Test
    void welcome() {
        try {
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.createPackage();
            MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
            mainDocumentPart.addStyledParagraphOfText("Title", "Hello World!");
            mainDocumentPart.addParagraphOfText("Welcome To Baeldung");
            File exportFile = new File("welcome.docx");
            wordPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template() {
        try {
            File doc = new File("Doc2.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage
                    .load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage
                    .getMainDocumentPart();
            String textNodesXPath = "//w:tbl";
            List<Object> textNodes = mainDocumentPart
                    .getJAXBNodesViaXPath(textNodesXPath, true);

            Tbl tbl = (Tbl) ((JAXBElement) textNodes.get(0)).getValue();
            Tr templateTr = (Tr) tbl.getContent().get(0);
            for (int i = 0; i < 5; i++) {
                Tr tr = XmlUtils.deepCopy(templateTr);
                tbl.getContent().add(tr);
            }

            Tbl cloneTbl = XmlUtils.deepCopy(tbl);
            mainDocumentPart.getContent().add(cloneTbl);

            File exportFile = new File("Docx2-updated.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template_2() {
        try {
            File doc = new File("template.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

            VariablePrepare.prepare(wordMLPackage);

            HashMap<String, String> variables = new HashMap<>();
            variables.put("firstName", "TEST");
            variables.put("lastName", "TEST");
            variables.put("salutation", "TEST");
            variables.put("message", "TEST");

            documentPart.variableReplace(variables);

            File exportFile = new File("template_2.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void contentControlsMergeXML() {
        try {
            // Without Saxon, you are restricted to XPath 1.0
            boolean USE_SAXON = true; // set this to true; add Saxon to your classpath, and uncomment below
//		if (USE_SAXON) XPathFactoryUtil.setxPathFactory(
//				new net.sf.saxon.xpath.XPathFactoryImpl());

            // the docx 'template'
            String input_DOCX = "binding-simple.docx";

            // the instance data
            String input_XML = "binding-simple-data.xml";

            // resulting docx
            String OUTPUT_DOCX = "OUT_ContentControlsMergeXML.docx";

            // Load input_template.docx
            WordprocessingMLPackage wordMLPackage = Docx4J.load(new File(input_DOCX));

            // Open the xml stream
            FileInputStream xmlStream = new FileInputStream(new File(input_XML));

            // Do the binding:
            // FLAG_NONE means that all the steps of the binding will be done,
            // otherwise you could pass a combination of the following flags:
            // FLAG_BIND_INSERT_XML: inject the passed XML into the document
            // FLAG_BIND_BIND_XML: bind the document and the xml (including any OpenDope handling)
            // FLAG_BIND_REMOVE_SDT: remove the content controls from the document (only the content remains)
            // FLAG_BIND_REMOVE_XML: remove the custom xml parts from the document

            //Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_NONE);
            //If a document doesn't include the Opendope definitions, eg. the XPathPart,
            //then the only thing you can do is insert the xml
            //the example document binding-simple.docx doesn't have an XPathPart....
            Docx4J.bind(wordMLPackage, xmlStream, Docx4J.FLAG_BIND_INSERT_XML | Docx4J.FLAG_BIND_BIND_XML);

            //Save the document
            Docx4J.save(wordMLPackage, new File(OUTPUT_DOCX), Docx4J.FLAG_NONE);
            System.out.println("Saved: " + OUTPUT_DOCX);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void test_placeHolderPattern() {
        try {
            Pattern placeHolderPattern = Pattern.compile("^\\$\\{([\\w\\d\\/]*)}");
            List<String> dataTerms = new ArrayList<>();

            File templateFile = new File("Doc3.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(templateFile);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            String textNodesXPath = "//w:t";
            List<Object> textNodes = mainDocumentPart.getJAXBNodesViaXPath(textNodesXPath, true);

            for (Object obj : textNodes) {
                Text text = (Text) ((JAXBElement) obj).getValue();
                String value = text.getValue();
                Matcher matcher = placeHolderPattern.matcher(value);
                System.out.println(value + " : " + matcher.matches());
                if (matcher.matches()) {
                    dataTerms.add(matcher.group(1));
                }
            }

            Map<String, String> dataDictionary = new HashMap<>();
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document doc = db.parse(new File("data.xml"));
            doc.getDocumentElement().normalize();
            XPath xPath = XPathFactory.newInstance().newXPath();

            for (String term : dataTerms) {
                NodeList queryResult = (NodeList) xPath.compile(term).evaluate(doc, XPathConstants.NODESET);
                if (queryResult.getLength() > 0) {
                    if (queryResult.getLength() == 1) {
                        dataDictionary.put(term, queryResult.item(0).getTextContent());
                    } else {
                        for (int i = 0; i < queryResult.getLength(); i++) {
                            dataDictionary.put(term + "/" + i, queryResult.item(i).getTextContent());
                        }
                    }
                }
            }

            System.out.println(dataDictionary);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void mergeData() {
        try {
            Pattern placeHolderPattern = Pattern.compile("\\{\\{([\\w\\d\\/]*)}}");

            List<String> dataTerms = new ArrayList<>();
            dataTerms.add("/company/staff/firstname");
            dataTerms.add("/company/staff/lastname");
            dataTerms.add("/company/staff/nickname");
            dataTerms.add("/company/staff/salary");

            Map<String, String> dataDictionary = new HashMap<>();
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document doc = db.parse(new File("data.xml"));
            doc.getDocumentElement().normalize();

            XPath xPath = XPathFactory.newInstance().newXPath();

            for (String term : dataTerms) {
                NodeList queryResult = (NodeList) xPath.compile(term).evaluate(doc, XPathConstants.NODESET);
                if (queryResult.getLength() > 0) {
                    if (queryResult.getLength() == 1) {
                        dataDictionary.put(term, queryResult.item(0).getTextContent());
                    } else {
                        for (int i = 0; i < queryResult.getLength(); i++) {
                            dataDictionary.put(term + "/" + i, queryResult.item(i).getTextContent());
                        }
                    }
                }
            }

            System.out.println(dataDictionary);

            File templateFile = new File("Doc3.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(templateFile);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            String textNodesXPath = "//w:t";
            List<Object> textNodes = mainDocumentPart
                    .getJAXBNodesViaXPath(textNodesXPath, true);

            for (Object obj : textNodes) {
                Text text = (Text) ((JAXBElement) obj).getValue();
                String placeHolder = text.getValue();
                System.out.println("Placeholder : " + placeHolder);
                if (dataDictionary.containsKey(placeHolder)) {
                    String value = dataDictionary.get(placeHolder);
                    System.out.println("Placeholder - value : " + value);
                    text.setValue(value);
                }
            }


            File exportFile = new File("Docx3-updated.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template_add_row() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
//        XPathFactoryUtil.getXPathFactory().set
        try {
            File doc = new File("Doc3.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
//            String trNodesXPath = "//w:t";
//            String trNodesXPath = "//w:t[matches(text(), \"{([\\w\\d/]+)\")]";
//            String trNodesXPath = "//w:t[fn:contains ( text(), \"name\")]";
//            String trNodesXPath = "//w:t[text() contains 'Lastname']";
//            String trNodesXPath = "//w:t[contains(text(), 'lastname')]";
            String trNodesXPath = "//w:t[matches(text(), '\\{([\\w\\d/]+)\\}')]/parent::w:r/parent::w:p/parent::w:tc/parent::w:tr";
            List<Object> trNodes = mainDocumentPart.getJAXBNodesViaXPath(trNodesXPath, false);

            for (Object obj : trNodes) {
                Tr templateRow = (Tr) obj;
                Tbl table = (Tbl) templateRow.getParent();
                int templateTrPos = table.getContent().indexOf(templateRow);
                for (int i = 0; i < 5; i++) {
                    Tr workingTr = XmlUtils.deepCopy(templateRow);
                    List<Object> texts = getAllElementFromObject(workingTr, Text.class);
                    for (Object obj1 : texts) {
                        Text text = (Text) obj1;
                        text.setValue(addRepeatParam(text.getValue(), i).orElse(text.getValue()));
                    }
                    table.getContent().add(templateTrPos + i + 1, workingTr);
                }
                table.getContent().remove(templateTrPos);
            }

            File exportFile = new File("Docx3-updated.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template_add_row_2() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
//        XPathFactoryUtil.getXPathFactory().set
        try {
            File doc = new File("Doc3.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
//            String trNodesXPath = "//w:t";
//            String trNodesXPath = "//w:t[matches(text(), \"{([\\w\\d/]+)\")]";
//            String trNodesXPath = "//w:t[fn:contains ( text(), \"name\")]";
//            String trNodesXPath = "//w:t[text() contains 'Lastname']";
//            String trNodesXPath = "//w:t[contains(text(), 'lastname')]";
            String repeatTRNodesXPath = "//w:t[matches(text(), '\\[([\\w\\d-]+)\\]')]/parent::w:r/parent::w:p/parent::w:tc/parent::w:tr";
            List<Object> trNodes = mainDocumentPart.getJAXBNodesViaXPath(repeatTRNodesXPath, false);

            for (Object obj : trNodes) {
                Tr templateRow = (Tr) obj;
                Tbl table = (Tbl) templateRow.getParent();
                int templateTrPos = table.getContent().indexOf(templateRow);
                for (int i = 0; i < 5; i++) {
                    Tr workingTr = XmlUtils.deepCopy(templateRow);
                    List<Object> texts = getAllElementFromObject(workingTr, Text.class);
                    for (Object obj1 : texts) {
                        Text text = (Text) obj1;
                        text.setValue(addRepeatParam(text.getValue(), i).orElse(text.getValue()));
                    }
                    table.getContent().add(templateTrPos + i + 1, workingTr);
                }
                table.getContent().remove(templateTrPos);
            }

            File exportFile = new File("Docx3-updated.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    }

    public Optional<String> addRepeatParam(String input, int i) {
        Pattern placeHolderPattern = Pattern.compile("^\\$\\{([\\w\\d\\/]*)}");
        Matcher matcher = placeHolderPattern.matcher(input);

        if (matcher.find()) {
            return Optional.of(input.replaceFirst("[\\w\\d\\/]+", matcher.group(1) + "/" + i));
        }
        return Optional.empty();
    }

    public Optional<String> addRepeatParam(String input) {
        Pattern placeHolderPattern = Pattern.compile("^\\$\\{([\\w\\d\\/]*)}");
        Matcher matcher = placeHolderPattern.matcher(input);

        if (matcher.find()) {
            return Optional.of(matcher.group(1));
        }
        return Optional.empty();
    }

    @Test
    void test_add_row() {
        String input = "${/company/staff/lastname}";
        Pattern placeHolderPattern = Pattern.compile("^\\$\\{([\\w\\d\\/]*)}");
        Matcher matcher = placeHolderPattern.matcher(input);

        if (matcher.find()) {
            System.out.println(input.replaceFirst("[\\w\\d\\/]+", matcher.group(1) + "/1"));
        }
    }

    private List<Object> cloneRepeatBlockContent(List<Object> contents) {
        List<Object> cloneContents = new ArrayList<>();
        for (Object content : contents) {
            Object cloneContent = XmlUtils.deepCopy(content);
            cloneContents.add(cloneContent);
        }
        return cloneContents;
    }

    private void fillData(Object workingContent, Document docData, String dataXpath, int index) {
        try {
            XPath xPath = XPathFactory.newInstance().newXPath();
            List<Object> texts = getAllElementFromObject(workingContent, Text.class);
            for (Object obj1 : texts) {
                Text text = (Text) obj1;
                Optional<String> dataField = addRepeatParam(text.getValue());
                if (dataField.isPresent()) {
                    if ("value".equals(dataField.get())) {
                        System.out.println(dataXpath);
                        NodeList queryResult = (NodeList) xPath.compile(dataXpath + "[" + index + "]")
                                .evaluate(docData, XPathConstants.NODESET);
                        if (queryResult.getLength() == 0) {
                            text.setValue("Not found");
                        } else {
                            text.setValue(queryResult.item(0).getTextContent());
                        }
                    } else {
                        System.out.println(dataXpath + "/" + dataField.get());
                        NodeList queryResult = (NodeList) xPath.compile(dataXpath + "[" + index + "]/"
                                + dataField.get()).evaluate(docData, XPathConstants.NODESET);
                        if (queryResult.getLength() == 0) {
                            text.setValue("Not found");
                        } else {
                            text.setValue(queryResult.item(0).getTextContent());
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void processRepeating(CTCustomXmlElement repeatBlock, String parentDataXpath, Document docData) {
        try {
            XPath xPath = XPathFactory.newInstance().newXPath();
            CTAttr dataAttr = repeatBlock.getCustomXmlPr().getAttr("data");
            String currentDataXpath = parentDataXpath + dataAttr.getVal();
            NodeList repeatCount = (NodeList) xPath.compile(currentDataXpath).evaluate(docData, XPathConstants.NODESET);
            List<Object> repeatContents = List.copyOf(repeatBlock.getContent());

            for (int i = 1; i <= repeatCount.getLength(); i++) {
                List<Object> workingContents = cloneRepeatBlockContent(repeatContents);
                for (Object obj : workingContents) {
                    if (obj instanceof JAXBElement) {
                        if (((JAXBElement) obj).getValue() instanceof CTCustomXmlElement) {
                            CTCustomXmlElement repeatSubBlock = (CTCustomXmlElement) ((JAXBElement) obj).getValue();
                            CTAttr repeatAttr = repeatBlock.getCustomXmlPr().getAttr("repeat");
                            if (Objects.isNull(repeatAttr) || !"true".equals(repeatAttr.getVal())) {
                                fillData(obj, docData, currentDataXpath, i);
                            } else {
                                processRepeating(repeatSubBlock, currentDataXpath + "[" + i + "]", docData);
                            }
                        }
                    } else {
                        fillData(obj, docData, currentDataXpath, i);
                    }
                }
                repeatBlock.getContent().addAll(workingContents);
            }
            repeatBlock.getContent().removeAll(repeatContents);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template_add_row_customxml_new() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document docData = db.parse(new File("data.xml"));
            docData.getDocumentElement().normalize();

            File doc = new File("Doc3-customxml.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

            /*Handle repeating*/
            String repeatPartXPath = "//w:attr[@w:name=\"loopLevel\"][@w:val=\"1\"]/parent::w:customXmlPr/parent::w:customXml";
            List<Object> repeatParts = mainDocumentPart.getJAXBNodesViaXPath(repeatPartXPath, false);

            for (Object obj : repeatParts) {
                CTCustomXmlElement repeatBlock = (CTCustomXmlElement) ((JAXBElement) obj).getValue();
                processRepeating(repeatBlock, "", docData);
            }

            File exportFile = new File("Docx3-customxml-updated.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template_add_row_customxml() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
//        XPathFactoryUtil.getXPathFactory().set
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document docData = db.parse(new File("data.xml"));
            docData.getDocumentElement().normalize();
            XPath xPath = XPathFactory.newInstance().newXPath();

            File doc = new File("Doc3-customxml.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();

            /*Handle display condition*/
//            String displayConditionXPath = "//w:attr[@w:name=\"displayCondition\"][@w:val=\"true\"]/parent::w:customXmlPr/parent::w:customXml";
//            List<Object> displayConditionParts = mainDocumentPart.getJAXBNodesViaXPath(displayConditionXPath, false);
//
//            for (Object obj : displayConditionParts) {
//                /* Hide */
//                Child child = (Child) ((JAXBElement) obj).getValue();
//                ContentAccessor parent = (ContentAccessor) child.getParent();
//                parent.getContent().remove(obj);
//            }

            /*Handle repeating*/
            String repeatPartXPath = "//w:attr[@w:name=\"repeat\"][@w:val=\"true\"]/parent::w:customXmlPr/parent::w:customXml";
            List<Object> repeatParts = mainDocumentPart.getJAXBNodesViaXPath(repeatPartXPath, false);

            for (Object obj : repeatParts) {
                CTCustomXmlElement repeatPart = (CTCustomXmlElement) ((JAXBElement) obj).getValue();
                CTAttr dataAttr = repeatPart.getCustomXmlPr().getAttr("data");
                NodeList repeatCount = (NodeList) xPath.compile(dataAttr.getVal()).evaluate(docData, XPathConstants.NODESET);
                List<Object> contents = List.copyOf(repeatPart.getContent());
                for (Object content : contents) {
                    for (int i = 1; i <= repeatCount.getLength(); i++) {
                        Object workingContent = XmlUtils.deepCopy(content);
                        List<Object> texts = getAllElementFromObject(workingContent, Text.class);
                        for (Object obj1 : texts) {
                            Text text = (Text) obj1;
                            Optional<String> dataField = addRepeatParam(text.getValue());
                            if (dataField.isPresent()) {
                                if ("value".equals(dataField.get())) {
                                    System.out.println(dataAttr.getVal());
                                    NodeList queryResult = (NodeList) xPath.compile(dataAttr.getVal() + "[" + i + "]")
                                            .evaluate(docData, XPathConstants.NODESET);
                                    if (queryResult.getLength() == 0) {
                                        text.setValue("Not found");
                                    } else {
                                        text.setValue(queryResult.item(0).getTextContent());
                                    }
                                } else {
                                    System.out.println(dataAttr.getVal() + "/" + dataField.get());
                                    NodeList queryResult = (NodeList) xPath.compile(dataAttr.getVal() + "[" + i + "]/"
                                            + dataField.get()).evaluate(docData, XPathConstants.NODESET);
                                    if (queryResult.getLength() == 0) {
                                        text.setValue("Not found");
                                    } else {
                                        text.setValue(queryResult.item(0).getTextContent());
                                    }
                                }
                            }
                        }
                        repeatPart.getContent().add(workingContent);
                    }
                    repeatPart.getContent().remove(content);
                }

            }

            File exportFile = new File("Docx3-customxml-updated.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void template_add_row_customxml_new_2() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document docData = db.parse(new File("data.xml"));
            docData.getDocumentElement().normalize();

            File doc = new File("Doc3-customxml.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            List<Object> contents = List.copyOf(mainDocumentPart.getContent());

            for (Object content : contents) {
                processContent(content, "", docData);
            }

            File exportFile = new File("Docx3-customxml-updated.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void processContent(Object content, String parentDataXpath, Document docData) {
        try {
            System.out.println(parentDataXpath);
            XPath xPath = XPathFactory.newInstance().newXPath();
            if (isContentControlElement(content)) {
                if (isCustomXmlElement(content)) {
                    CTCustomXmlElement customXmlElement = (CTCustomXmlElement) ((JAXBElement) content).getValue();
                    CTAttr dataAttr = customXmlElement.getCustomXmlPr().getAttr("data");
                    if (isRepeat(customXmlElement)) {
                        System.out.println(parentDataXpath + dataAttr.getVal());
                        NodeList repeatCount = (NodeList) xPath.compile(parentDataXpath + dataAttr.getVal()).evaluate(docData, XPathConstants.NODESET);
                        List<Object> repeatContents = List.copyOf(customXmlElement.getContent());
                        for (int i = 1; i <= repeatCount.getLength(); i++) {
                            List<Object> workingContents = cloneRepeatBlockContent(repeatContents);
                            for (Object obj : workingContents) {
                                processContent(obj, parentDataXpath + dataAttr.getVal() + "[" + i + "]", docData);
                            }
                            customXmlElement.getContent().addAll(workingContents);
                        }
                        customXmlElement.getContent().removeAll(repeatContents);
                    } else {
                        String dataXpath = "";

                        if (Objects.nonNull(dataAttr)) {
                            dataXpath = dataAttr.getVal();
                        }

                        for (Object childContent : customXmlElement.getContent()) {
                            processContent(childContent, parentDataXpath + dataXpath, docData);
                        }
                    }
                } else {
                    ContentAccessor contentAccessor;

                    if (content instanceof JAXBElement) {
                        contentAccessor = (ContentAccessor) ((JAXBElement) content).getValue();
                    } else {
                        contentAccessor = (ContentAccessor) content;
                    }

                    for (Object childContent : contentAccessor.getContent()) {
                        processContent(childContent, parentDataXpath, docData);
                    }
                }
            } else {
                List<Object> texts = getAllElementFromObject(content, Text.class);
                for (Object obj1 : texts) {
                    Text text = (Text) obj1;
                    Optional<String> dataField = addRepeatParam(text.getValue());
                    if (dataField.isPresent()) {
                        if ("value".equals(dataField.get())) {
                            NodeList queryResult = (NodeList) xPath.compile(parentDataXpath)
                                    .evaluate(docData, XPathConstants.NODESET);
                            if (queryResult.getLength() == 0) {
                                text.setValue("Not found");
                            } else {
                                text.setValue(queryResult.item(0).getTextContent());
                            }
                        } else {
                            System.out.println(parentDataXpath + "/" + dataField.get());
                            NodeList queryResult = (NodeList) xPath.compile(parentDataXpath + "/"
                                    + dataField.get()).evaluate(docData, XPathConstants.NODESET);
                            if (queryResult.getLength() == 0) {
                                text.setValue("Not found");
                            } else {
                                text.setValue(queryResult.item(0).getTextContent());
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private boolean isRepeat(CTCustomXmlElement customXmlElement) {
        CTAttr repeatAttr = customXmlElement.getCustomXmlPr().getAttr("repeat");
        if (Objects.nonNull(repeatAttr) && "true".equals(repeatAttr.getVal())) {
            return true;
        }
        return false;
    }

    private boolean isCustomXmlElement(Object element) {
        if (element instanceof JAXBElement) {
            element = ((JAXBElement<?>) element).getValue();
        }

        if (element instanceof CTCustomXmlElement) {
            return true;
        }
        return false;
    }

    private boolean isContentControlElement(Object content) {
        return isContentControlElement(content, CTCustomXmlRow.class)
                || isContentControlElement(content, CTCustomXmlBlock.class);
    }

    public boolean isContentControlElement(Object obj, Class<?> toSearch) {
        if (obj instanceof JAXBElement) {
            obj = ((JAXBElement<?>) obj).getValue();
        }

        if (obj.getClass().equals(toSearch)) {
            return true;
        } else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                if (isContentControlElement(child, toSearch)) {
                    return true;
                }
            }
        }

        return false;
    }
}
