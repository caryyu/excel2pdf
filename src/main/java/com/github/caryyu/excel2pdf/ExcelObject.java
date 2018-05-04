package com.github.caryyu.excel2pdf;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.*;

import com.itextpdf.text.*;

/**
 * Created by cary on 6/15/17.
 */
public class ExcelObject {
    /**
     * 锚名称
     */
    private String anchorName;
    /**
     * Excel Stream
     */
    private InputStream inputStream;
    /**
     * POI Excel
     */
    private Excel excel;

    public ExcelObject(InputStream inputStream){
        this.inputStream = inputStream;
        this.excel = new Excel(this.inputStream);
    }

    public ExcelObject(String anchorName , InputStream inputStream){
        this.anchorName = anchorName;
        this.inputStream = inputStream;
        this.excel = new Excel(this.inputStream);
    }
    public String getAnchorName() {
        return anchorName;
    }
    public void setAnchorName(String anchorName) {
        this.anchorName = anchorName;
    }
    public InputStream getInputStream() {
        return this.inputStream;
    }
    public void setInputStream(InputStream inputStream) {
        this.inputStream = inputStream;
    }
    Excel getExcel() {
        return excel;
    }
    
    /**
     * 取得 Excel 版面設定的紙張大小並依據直橫式旋轉
     * @return
     */
    public Rectangle getPageSize() {
    	short size = excel.getSheet().getPrintSetup().getPaperSize();
    	boolean landScape = excel.getSheet().getPrintSetup().getLandscape();
    	switch(size) {
    	case PrintSetup.A3_PAPERSIZE:
    		return landScape?PageSize.A3.rotate():PageSize.A3;
    	case PrintSetup.A4_PAPERSIZE:
    		return landScape?PageSize.A4.rotate():PageSize.A4;
    	case PrintSetup.A4_ROTATED_PAPERSIZE:
        	return landScape?PageSize.A4:PageSize.A4.rotate();
        case PrintSetup.A4_SMALL_PAPERSIZE:
           	return landScape?PageSize.A4.rotate():PageSize.A4;
        case PrintSetup.A5_PAPERSIZE:
           	return landScape?PageSize.A5.rotate():PageSize.A5;
        case PrintSetup.LETTER_PAPERSIZE:
    		return landScape?PageSize.LETTER.rotate():PageSize.LETTER;
        case PrintSetup.LETTER_ROTATED_PAPERSIZE:
    		return landScape?PageSize.LETTER:PageSize.LETTER.rotate();
        case PrintSetup.B4_PAPERSIZE:
           	return landScape?PageSize.B4.rotate():PageSize.B4;
        case PrintSetup.B5_PAPERSIZE:
           	return landScape?PageSize.B5.rotate():PageSize.B5;
    	default: return landScape?PageSize.A4.rotate().rotate():PageSize.A4;
    	}
    }
}