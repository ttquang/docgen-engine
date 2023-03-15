package com.example.docx4jexample1;

import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.InputStream;

public class DataSourceFactory {
    public static DataSource constructXMLDataSource(InputStream inputStream) {
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document document = db.parse(inputStream);
            document.getDocumentElement().normalize();

            return new XMLDataSource(document);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public static DataSource constructXMLDataSource(File file) {
        try {
            DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
            DocumentBuilder db = dbf.newDocumentBuilder();
            Document document = db.parse(file);
            document.getDocumentElement().normalize();

            return new XMLDataSource(document);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}
