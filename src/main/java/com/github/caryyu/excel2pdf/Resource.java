package com.github.caryyu.excel2pdf;

import org.apache.poi.hssf.usermodel.*;

import com.itextpdf.text.*;
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
            // 搜尋系統,載入系統內的字型(慢)
            FontFactory.registerDirectories();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 將 POI Font 轉換到 iText Font
     * @param font
     * @return
     */
    public static com.itextpdf.text.Font getFont(HSSFFont font) {
        try {
            com.itextpdf.text.Font iTextFont = FontFactory.getFont(font.getFontName(),
                    BaseFont.IDENTITY_H, BaseFont.EMBEDDED,
                    font.getFontHeightInPoints());
            return iTextFont;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }
}