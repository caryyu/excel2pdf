package com.github.caryyu.excel2pdf;

import com.itextpdf.text.DocumentException;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Excel2PDFTest {
    @Test
    public void testOrigin() throws IOException, DocumentException {
        FileInputStream fis1 = new FileInputStream(new File("D:\\pdfexport\\MAD 5-3-05-Octavia NF-20131025.xls"));
        FileInputStream fis2 = new FileInputStream(new File("D:\\pdfexport\\MAD 6-1-47-Octavia NF-20131025.xls"));
        FileInputStream fis3 = new FileInputStream(new File("D:\\pdfexport\\MAD 038-Superb FL DS-20131025.xls"));

        FileOutputStream fos = new FileOutputStream(new File("D:\\test.pdf"));

        List<ExcelObject> objects = new ArrayList<ExcelObject>();
        objects.add(new ExcelObject("1.MAD 5-3-05-Octavia NF-20131025.xls",fis1));
        objects.add(new ExcelObject("2.MAD 6-1-47-Octavia NF-20131025.xls",fis2));
        objects.add(new ExcelObject("3.MAD 038-Superb FL DS-20131025.xls",fis3));

        Excel2Pdf pdf = new Excel2Pdf(objects , fos);

        pdf.convert();
    }
}
