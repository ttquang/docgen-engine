package com.example.docx4jexample1;

public enum SdtBlockType {
    TEXT("Text"),
    LOOP("Loop"),
    IF("If");

    public final String label;

    SdtBlockType(String label) {
        this.label = label;
    }

    public static SdtBlockType valueOfLabel(String label) {
        for (SdtBlockType e : values()) {
            if (e.label.equals(label)) {
                return e;
            }
        }
        return null;
    }
}
