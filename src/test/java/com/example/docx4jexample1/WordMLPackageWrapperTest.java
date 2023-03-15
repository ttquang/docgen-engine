package com.example.docx4jexample1;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import org.docx4j.utils.XPathFactoryUtil;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.*;

@SpringBootTest
class WordMLPackageWrapperTest {

    @Test
    void contextLoads() {
    }

    @Test
    void template_sdt() {
        XPathFactoryUtil.setxPathFactory(new net.sf.saxon.xpath.XPathFactoryImpl());
        try {
            WordMLPackageWrapper wordMLPackageWrapper =
                    new WordMLPackageWrapper(new File("Cheque_Issuance_Form.docx"),new File("Cheque_Issuance_Form-data.xml"));
            wordMLPackageWrapper.process();
            wordMLPackageWrapper.save(new File("Cheque_Issuance_Form-result.docx"));


//            document4j
            File inputWord = new File("Cheque_Issuance_Form-result.docx");
            File outputFile = new File("Cheque_Issuance_Form-result.pdf");
            try  {
                InputStream docxInputStream = new FileInputStream(inputWord);
                OutputStream outputStream = new FileOutputStream(outputFile);
                IConverter converter = LocalConverter.builder().build();
                converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
                outputStream.close();
                System.out.println("success");
            } catch (Exception e) {
                e.printStackTrace();
            }

//            try {
//                WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new java.io.File("tmb-docgen-1-result.docx"));
//                FileOutputStream os = new FileOutputStream(new File("tmb-docgen-1-result.pdf"));
//
//                FieldUpdater updater = null;
//                Mapper fontMapper = new IdentityPlusMapper();
//                wordMLPackage.setFontMapper(fontMapper);
//                PhysicalFont font = PhysicalFonts.get("Cordia New");
//                FOSettings foSettings = new FOSettings(wordMLPackage);
//                foSettings.setFoDumpFile(new java.io.File("tmb-docgen-1-result.fo"));
//                FopFactoryBuilder fopFactoryBuilder = FORendererApacheFOP.getFopFactoryBuilder(foSettings);
//                FopFactory fopFactory = fopFactoryBuilder.build();
//                FOUserAgent foUserAgent = FORendererApacheFOP.getFOUserAgent(foSettings, fopFactory);
//                // configure foUserAgent as desired
//                foUserAgent.setTitle("my title");
//                Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
//                updater = null;
//                foSettings = null;
//                wordMLPackage = null;
//            } catch (Exception e) {
//                e.printStackTrace();
//            }
//            Document doc = new Document("tmb-docgen-1-result.docx");
//            //Save the document to PDF
//            doc.saveToFile("tmb-docgen-1-result.pdf", FileFormat.PDF);
//            doc.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
