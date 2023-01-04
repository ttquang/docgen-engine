package com.example.docx4jexample1;

public enum SdtType {
    TEXT("Text"),
    LOOP("Loop"),
    IF("If");

    public final String label;

    SdtType(String label) {
        this.label = label;
    }

    public static SdtType valueOfLabel(String label) {
        for (SdtType e : values()) {
            if (e.label.equals(label)) {
                return e;
            }
        }
        return null;
    }
}
