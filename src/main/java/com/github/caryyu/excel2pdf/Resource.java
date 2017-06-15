package com.github.caryyu.excel2pdf;

import com.itextpdf.text.pdf.BaseFont;

/**
 * Created by cary on 6/15/17.
 */
public class Resource {
    /**
     * 中文字体支持
     */
    protected static BaseFont BASE_FONT_CHINESE;
    static {
        try {
            BASE_FONT_CHINESE = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}