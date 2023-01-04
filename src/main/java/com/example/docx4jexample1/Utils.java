package com.example.docx4jexample1;

import org.docx4j.jaxb.Context;
import org.docx4j.wml.*;

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

            for(int var9 = 0; var9 < var8; ++var9) {
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
}
