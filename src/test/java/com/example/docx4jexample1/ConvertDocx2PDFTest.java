package com.example.docx4jexample1;

import com.aspose.words.Document;
import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.convert.out.pdf.PdfConversion;
import org.docx4j.convert.out.pdf.viaXSLFO.PdfSettings;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;

@SpringBootTest
class ConvertDocx2PDFTest {

    @Test
    void contextLoads() {
    }

    @Test
    void test_convert_Apache_POI() {
        File inputWord = new File("SP_Old_staff_and_TMBCare_Topup_Agreement_without_Duty_Stamp.docx");
        File outputFile = new File("SP_Old_staff_and_TMBCare_Topup_Agreement_without_Duty_Stamp-ApachePOI.pdf");
        try {
            InputStream is = new FileInputStream(inputWord);
            OutputStream out = new FileOutputStream(outputFile);
            XWPFDocument document = new XWPFDocument(is);
            PdfOptions options = PdfOptions.create();
            PdfConverter.getInstance().convert(document, out, options);
        } catch (Throwable e) {
            e.printStackTrace();
        }
    }

    @Test
    void test_convert_document4j() {
        File inputWord = new File("CC_Application_Form-result.docx");
        File outputFile = new File("CC_Application_Form-result-document4j.pdf");
        try {
            InputStream docxInputStream = new FileInputStream(inputWord);
            OutputStream outputStream = new FileOutputStream(outputFile);
            IConverter converter = LocalConverter.builder().build();
            converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void test_convert_Aspose() {
        File inputWord = new File("SP_Old_staff_and_TMBCare_Topup_Agreement_without_Duty_Stamp.docx");
        File outputFile = new File("SP_Old_staff_and_TMBCare_Topup_Agreement_without_Duty_Stamp-Aspose.pdf");
        try {
            // Load the Word document from disk
            Document doc = new Document("SP_Old_staff_and_TMBCare_Topup_Agreement_without_Duty_Stamp.docx");
            // Save as PDF
            doc.save("SP_Old_staff_and_TMBCare_Topup_Agreement_without_Duty_Stamp-Aspose.pdf");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Test
    void test_convert_docx4j() {
        File inputWord = new File("SP_Old_staff_and_TMBCare_Topup_Agreement_without_Duty_Stamp.docx");
        File outputFile = new File("SP_Old_staff_and_TMBCare_Topup_Agreement_without_Duty_Stamp-docx4j.pdf");
        try {
            long start = System.currentTimeMillis();

            // 1) Load DOCX into WordprocessingMLPackage
            InputStream is = new FileInputStream(inputWord);
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(is);

            // 2) Prepare Pdf settings
            PdfSettings pdfSettings = new PdfSettings();

            // 3) Convert WordprocessingMLPackage to Pdf
            OutputStream out = new FileOutputStream(outputFile);
            PdfConversion converter = new org.docx4j.convert.out.pdf.viaXSLFO.Conversion(
                    wordMLPackage);
            converter.output(out, pdfSettings);

            System.err.println("Generate pdf/HelloWorld.pdf with "
                    + (System.currentTimeMillis() - start) + "ms");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
