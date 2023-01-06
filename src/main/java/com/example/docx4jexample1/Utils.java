package com.example.docx4jexample1;

import jakarta.xml.bind.JAXBElement;
import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Utils {
    public static P createParagraphOfText(String simpleText) {
        ObjectFactory factory = Context.getWmlObjectFactory();
        P para = factory.createP();
        if (simpleText != null) {
            String[] splits = simpleText.replaceAll("\r", "").split("\n");
            boolean afterNewline = false;
            R run = factory.createR();
            String[] var7 = splits;
            int var8 = splits.length;

            for (int var9 = 0; var9 < var8; ++var9) {
                String s = var7[var9];
                if (afterNewline) {
                    Br br = factory.createBr();
                    run.getContent().add(br);
                }

                Text t = factory.createText();
                t.setValue(s);
                if (s.startsWith(" ")) {
                    t.setSpace("preserve");
                } else if (s.endsWith(" ")) {
                    t.setSpace("preserve");
                }

                run.getContent().add(t);
                afterNewline = true;
            }

            para.getContent().add(run);
        }

        return para;
    }

    public static boolean isContentControl(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    if ("Loop".equals(alias.getVal()) || "Text".equals(alias.getVal()) || alias.getVal().startsWith("If"))
                        return true;
                }
            }
        }
        return false;
    }

    public static boolean isTextControl(SdtElement sdtElement) {
        for (Object property : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (property instanceof JAXBElement) {
                property = ((JAXBElement) property).getValue();
            }

            if (property instanceof SdtPr.Alias) {
                SdtPr.Alias alias = (SdtPr.Alias) property;
                if ("Text".equals(alias.getVal()))
                    return true;
            }
        }
        return false;
    }

    public static boolean isLoopControl(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    if ("Loop".equals(alias.getVal()))
                        return true;
                }
            }
        }
        return false;
    }

    public static boolean isIfControl(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    if (alias.getVal().startsWith("If"))
                        return true;
                }
            }
        }
        return false;
    }

    public static String getTag(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof Tag) {
                return ((Tag) prObj).getVal();
            }
        }
        return "";
    }

    public static String getAlias(SdtElement sdtElement) {
        for (Object prObj : sdtElement.getSdtPr().getRPrOrAliasOrLock()) {
            if (prObj instanceof JAXBElement) {
                if (((JAXBElement) prObj).getValue() instanceof SdtPr.Alias) {
                    SdtPr.Alias alias = (SdtPr.Alias) ((JAXBElement) prObj).getValue();
                    return alias.getVal();
                }
            }
        }
        throw new IllegalStateException();
    }

    public static String getIfExpression(String input) {
        Pattern ifExpressPattern = Pattern.compile("^If\\(([\\w\\d =']+)\\)");
        Matcher matcher = ifExpressPattern.matcher(input);
        if (matcher.matches()) {
            return matcher.group(1);
        }
        return "";
    }

}
