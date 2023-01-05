package com.example.docx4jexample1;

import org.docx4j.dml.chart.CTBarChart;
import org.docx4j.dml.chart.CTBarSer;
import org.docx4j.dml.chart.CTNumDataSource;
import org.docx4j.dml.chart.CTNumVal;
import org.docx4j.openpackaging.contenttype.ContentType;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.utils.XPathFactoryUtil;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.util.List;

@SpringBootTest
class ChartTests {

    @Test
    void contextLoads() {
    }

    @Test
    void template_sdt() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
        try {
            File doc = new File("Chart.docx");
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(doc);
            MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
            RelationshipsPart relationshipsPart = mainDocumentPart.getRelationshipsPart();
            Relationship relationship = relationshipsPart.getRelationshipByType(Namespaces.SPREADSHEETML_CHART);
            Chart chart = (Chart) relationshipsPart.getPart(relationship);

            CTBarChart barChart = (CTBarChart) chart.getContents().getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart().get(0);
            List<CTBarSer> serList = barChart.getSer();
            CTNumDataSource serVal = serList.get(0).getVal();
            List<CTNumVal> ptList = serVal.getNumRef().getNumCache().getPt();
            for (CTNumVal pt : ptList) {
                System.out.println(pt.getV());
                pt.setV("15");
            }

            File exportFile = new File("Chart-result.docx");
            wordMLPackage.save(exportFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void chartSample() {
//        try {
//            ContentType contentType = new ContentType(
//                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//            String excelFilePath = "e://data.xlsx";
//            PresentationMLPackage presentationMLPackage = null;
//            int slideIndex = 11;
//            String templateFile = "/main/resources/template.pptx";
//            String destinationFolderStr = "c:/ppt/";
//            File destinationFolder = new File(destinationFolderStr);
//            if (!destinationFolder.exists()) {
//                destinationFolder.mkdir();
//            }
//            destinationFolderStr = destinationFolderStr + File.separator
//                    + "docx4JChart.pptx";
//            try {
//                InputStream in = Docx4jChart.class
//                        .getResourceAsStream(templateFile);
//                presentationMLPackage = (PresentationMLPackage) OpcPackage
//                        .load(in);
//            } catch (Exception e) {
//                // TODO: handle exception
//                presentationMLPackage = PresentationMLPackage.createPackage();
//            }
//
//            MainPresentationPart pp = (MainPresentationPart) presentationMLPackage
//                    .getParts().getParts()
//                    .get(new PartName("/ppt/presentation.xml"));
//            SlideLayoutPart layoutPart = (SlideLayoutPart) presentationMLPackage
//                    .getParts().getParts()
//                    .get(new PartName("/ppt/slideLayouts/slideLayout2.xml"));
//
//            SlidePart slide1 = presentationMLPackage.createSlidePart(pp,
//                    layoutPart, new PartName("/ppt/slides/slide" + slideIndex
//                            + ".xml"));
//            StringBuffer slideXMLBuffer = new StringBuffer();
//            BufferedReader br = null;
//            String line = "";
//            String slideDataXmlFile = "/main/data/slide_data.xml";
//            InputStream in = Docx4jChart.class
//                    .getResourceAsStream(slideDataXmlFile);
//            Reader fr = new InputStreamReader(in, "utf-8");
//            br = new BufferedReader(fr);
//            while ((line = br.readLine()) != null) {
//                slideXMLBuffer.append(line);
//                slideXMLBuffer.append(" ");
//            }
//
//            Sld sld = (Sld) XmlUtils.unmarshalString(slideXMLBuffer.toString(),
//                    Context.jcPML, Sld.class);
//            slide1.setJaxbElement(sld);
//
//            Chart chartPart = new Chart(
//                    new PartName("/ppt/charts/chart" + slideIndex + ".xml"));
//
//            StringBuffer chartXMLBuffer = new StringBuffer();
//            String chartDataXmlFile = "/main/data/chart_data.xml";
//            in = Docx4jChart.class.getResourceAsStream(chartDataXmlFile);
//            fr = new InputStreamReader(in, "utf-8");
//            br = new BufferedReader(fr);
//            while ((line = br.readLine()) != null) {
//                chartXMLBuffer.append(line);
//                chartXMLBuffer.append(" ");
//            }
//
//            CTChartSpace chartSpace = (CTChartSpace) XmlUtils.unmarshalString(
//                    chartXMLBuffer.toString(), Context.jcPML,
//                    CTChartSpace.class);
//            chartPart.setJaxbElement(chartSpace);
//
//            slide1.addTargetPart(chartPart);
//            EmbeddedPackagePart embeddedPackagePart = new EmbeddedPackagePart(
//                    new PartName("/ppt/embeddings/Microsoft_Excel_Worksheet"
//                            + slideIndex + ".xlsx"));
//            embeddedPackagePart.setContentType(contentType);
//            embeddedPackagePart
//                    .setRelationshipType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
//            RelationshipsPart owningRelationshipPart = new RelationshipsPart();
//            PartName partName = new PartName("/ppt/charts/_rels/chart"
//                    + slideIndex + ".xml.rels");
//            owningRelationshipPart.setPartName(partName);
//            Relationship relationship = new Relationship();
//            relationship.setId("rId1");
//            relationship.setTarget("../embeddings/Microsoft_Excel_Worksheet"
//                    + slideIndex + ".xlsx");
//            relationship
//                    .setType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package");
//            owningRelationshipPart.setRelationships(new Relationships());
//            owningRelationshipPart.addRelationship(relationship);
//            embeddedPackagePart
//                    .setOwningRelationshipPart(owningRelationshipPart);
//            embeddedPackagePart.setPackage(presentationMLPackage);
//            String dataFile = "/main/data/data.xlsx";
//            InputStream inputStream = Docx4jChart.class
//                    .getResourceAsStream(dataFile);
//            embeddedPackagePart.setBinaryData(inputStream);
//            chartPart.addTargetPart(embeddedPackagePart);
//            presentationMLPackage.save(new java.io.File(destinationFolderStr));
//            System.out.println("Powerpoint Office document is created in the following path: "+destinationFolderStr);
//        } catch (Exception exception) {
//            exception.printStackTrace();
//        }
    }

}
