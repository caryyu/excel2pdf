package com.github.caryyu.excel2pdf;

import com.itextpdf.text.Document;

import java.io.OutputStream;

/**
 * Created by cary on 6/15/17.
 */
public class PdfTool {
    //
    protected Document document;
    //
    protected OutputStream os;

    public Document getDocument() {
        if (document == null) {
            document = new Document();
        }
        return document;
    }
}