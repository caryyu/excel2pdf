package com.github.caryyu.excel2pdf;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Sheet;

import java.awt.Dimension;
import java.util.List;

/**
 * Created by cary on 6/15/17.
 */
public class POIImage {
    protected Dimension dimension;
    protected byte[] bytes;
    protected ClientAnchor anchor;

    public POIImage getCellImage(Cell cell) {
        Sheet sheet = cell.getSheet();
        if (sheet instanceof HSSFSheet) {
            HSSFSheet hssfSheet = (HSSFSheet) sheet;
            if (hssfSheet.getDrawingPatriarch() != null) {
                List<HSSFShape> shapes = hssfSheet.getDrawingPatriarch().getChildren();
                for (HSSFShape shape : shapes) {
                    HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
                    if (shape instanceof HSSFPicture) {
                        HSSFPicture pic = (HSSFPicture) shape;
                        PictureData data = pic.getPictureData();
                        String extension = data.suggestFileExtension();
                        int row1 = anchor.getRow1();
                        int row2 = anchor.getRow2();
                        int col1 = anchor.getCol1();
                        int col2 = anchor.getCol2();
                        if (row1 == cell.getRowIndex() && col1 == cell.getColumnIndex()) {
                            dimension = pic.getImageDimension();
                            this.anchor = anchor;
                            this.bytes = data.getData();
                        }
                    }
                }
            }
        }
        return this;
    }

    public Dimension getDimension() {
        return dimension;
    }

    public void setDimension(Dimension dimension) {
        this.dimension = dimension;
    }

    public byte[] getBytes() {
        return bytes;
    }

    public void setBytes(byte[] bytes) {
        this.bytes = bytes;
    }

    public ClientAnchor getAnchor() {
        return anchor;
    }

    public void setAnchor(ClientAnchor anchor) {
        this.anchor = anchor;
    }
}