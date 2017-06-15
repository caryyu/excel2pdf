package com.github.caryyu.excel2pdf;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by cary on 6/15/17.
 */
public class PdfTableExcel {
    //ExcelObject
    protected ExcelObject excelObject;
    //excel
    protected Excel excel;
    //
    protected boolean setting = false;

    /**
     * <p>Description: Constructor</p>
     * @param excelObject
     */
    public PdfTableExcel(ExcelObject excelObject){
        this.excelObject = excelObject;
        this.excel = excelObject.getExcel();
    }

    /**
     * <p>Description: 获取转换过的Excel内容Table</p>
     * @return PdfPTable
     * @throws BadElementException
     * @throws MalformedURLException
     * @throws IOException
     */
    public PdfPTable getTable() throws BadElementException, MalformedURLException, IOException {
        Sheet sheet = this.excel.getSheet();
        return toParseContent(sheet);
    }

    protected PdfPTable toParseContent(Sheet sheet) throws BadElementException, MalformedURLException, IOException{
        int rowlength = sheet.getLastRowNum();
        List<PdfPCell> cells = new ArrayList<PdfPCell>();
        float[] widths = null;
        float mw = 0;
        for (int i = 0; i < rowlength; i++) {
            Row row = sheet.getRow(i);
            float[] cws = new float[row.getLastCellNum()];
            for (int j = 0; j < row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                float cw = getPOIColumnWidth(cell);
                cws[cell.getColumnIndex()] = cw;
                if(isUsed(cell.getColumnIndex(), row.getRowNum())){
                    continue;
                }
                cell.setCellType(Cell.CELL_TYPE_STRING);
                CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex());
                //
                int rowspan = 1;
                int colspan = 1;
                if (range != null) {
                    rowspan = range.getLastRow() - range.getFirstRow() + 1;
                    colspan = range.getLastColumn() - range.getFirstColumn() + 1;
                }
                //PDF单元格
                PdfPCell pdfpCell = new PdfPCell();
                pdfpCell.setBackgroundColor(new BaseColor(getBackgroundColorByExcel(cell.getCellStyle())));
                pdfpCell.setColspan(colspan);
                pdfpCell.setRowspan(rowspan);
                pdfpCell.setVerticalAlignment(getVAlignByExcel(cell.getCellStyle().getVerticalAlignment()));
                pdfpCell.setHorizontalAlignment(getHAlignByExcel(cell.getCellStyle().getAlignment()));
                pdfpCell.setPhrase(getPhrase(cell));
                pdfpCell.setFixedHeight(this.getPixelHeight(row.getHeightInPoints()));
                addBorderByExcel(pdfpCell, cell.getCellStyle());
                addImageByPOICell(pdfpCell , cell , cw);
                //
                cells.add(pdfpCell);
                j += colspan - 1;
            }
            float rw = 0;
            for (int j = 0; j < cws.length; j++) {
                rw += cws[j];
            }
            if (rw > mw ||  mw == 0) {
                widths = cws;
                mw = rw;
            }
        }
        //
        PdfPTable table = new PdfPTable(widths);
        table.setWidthPercentage(100);
//        table.setLockedWidth(true);
        for (PdfPCell pdfpCell : cells) {
            table.addCell(pdfpCell);
        }
        return table;
    }

    protected Phrase getPhrase(Cell cell){
        if(this.setting || this.excelObject.getAnchorName() == null){
            return new Phrase(cell.getStringCellValue(), getFontByExcel(cell.getCellStyle()));
        }
        Anchor anchor = new Anchor(cell.getStringCellValue() , getFontByExcel(cell.getCellStyle()));
        anchor.setName(this.excelObject.getAnchorName());
        this.setting = true;
        return anchor;
    }

    protected void addImageByPOICell(PdfPCell pdfpCell , Cell cell , float cellWidth) throws BadElementException, MalformedURLException, IOException{
        POIImage poiImage = new POIImage().getCellImage(cell);
        byte[] bytes = poiImage.getBytes();
        if(bytes != null){
//           double cw = cellWidth;
//           double ch = pdfpCell.getFixedHeight();
//
//           double iw = poiImage.getDimension().getWidth();
//           double ih = poiImage.getDimension().getHeight();
//
//           double scale = cw / ch;
//
//           double nw = iw * scale;
//           double nh = ih - (iw - nw);
//
//           POIUtil.scale(bytes , nw  , nh);
            pdfpCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
            pdfpCell.setHorizontalAlignment(Element.ALIGN_CENTER);
            Image image = Image.getInstance(bytes);
            pdfpCell.setImage(image);
        }
    }

    protected float getPixelHeight(float poiHeight){
        float pixel = poiHeight / 28.6f * 26f;
        return pixel;
    }

    /**
     * <p>Description: 此处获取Excel的列宽像素(无法精确实现,期待有能力的朋友进行改善此处)</p>
     * @param cell
     * @return 像素宽
     */
    protected int getPOIColumnWidth(Cell cell) {
        int poiCWidth = excel.getSheet().getColumnWidth(cell.getColumnIndex());
        int colWidthpoi = poiCWidth;
        int widthPixel = 0;
        if (colWidthpoi >= 416) {
            widthPixel = (int) (((colWidthpoi - 416.0) / 256.0) * 8.0 + 13.0 + 0.5);
        } else {
            widthPixel = (int) (colWidthpoi / 416.0 * 13.0 + 0.5);
        }
        return widthPixel;
    }

    protected CellRangeAddress getColspanRowspanByExcel(int rowIndex, int colIndex) {
        CellRangeAddress result = null;
        Sheet sheet = excel.getSheet();
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstColumn() == colIndex && range.getFirstRow() == rowIndex) {
                result = range;
            }
        }
        return result;
    }

    protected boolean isUsed(int colIndex , int rowIndex){
        boolean result = false;
        Sheet sheet = excel.getSheet();
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            if (firstRow < rowIndex && lastRow >= rowIndex) {
                if(firstColumn <= colIndex && lastColumn >= colIndex){
                    result = true;
                }
            }
        }
        return result;
    }

    protected Font getFontByExcel(CellStyle style) {
        Font result = new Font(Resource.BASE_FONT_CHINESE , 8 , Font.NORMAL);
        Workbook wb = excel.getWorkbook();
        //字体样式索引
        short index = style.getFontIndex();
        org.apache.poi.ss.usermodel.Font font = wb.getFontAt(index);
        //字体颜色
        int colorIndex = font.getColor();
        if(font.getBoldweight() == org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD){
            result.setStyle(Font.BOLD);
        }
        HSSFColor color = HSSFColor.getIndexHash().get(colorIndex);
        if(color != null){
            int rbg = POIUtil.getRGB(color);
            result.setColor(new BaseColor(rbg));
        }
        //下划线
        FontUnderline underline = FontUnderline.valueOf(font.getUnderline());
        if(underline == FontUnderline.SINGLE){
            String ulString = Font.FontStyle.UNDERLINE.getValue();
            result.setStyle(ulString);
        }
        return result;
    }

    protected int getBackgroundColorByExcel(CellStyle style) {
        Color color = style.getFillForegroundColorColor();
        return POIUtil.getRGB(color);
    }

    protected void addBorderByExcel(PdfPCell cell , CellStyle style) {
        Workbook wb = excel.getWorkbook();
        cell.setBorderColorLeft(new BaseColor(POIUtil.getBorderRBG(wb,style.getLeftBorderColor())));
        cell.setBorderColorRight(new BaseColor(POIUtil.getBorderRBG(wb,style.getRightBorderColor())));
        cell.setBorderColorTop(new BaseColor(POIUtil.getBorderRBG(wb,style.getTopBorderColor())));
        cell.setBorderColorBottom(new BaseColor(POIUtil.getBorderRBG(wb,style.getBottomBorderColor())));
    }

    protected int getVAlignByExcel(short align) {
        int result = 0;
        if (align == CellStyle.VERTICAL_BOTTOM) {
            result = Element.ALIGN_BOTTOM;
        }
        if (align == CellStyle.VERTICAL_CENTER) {
            result = Element.ALIGN_MIDDLE;
        }
        if (align == CellStyle.VERTICAL_JUSTIFY) {
            result = Element.ALIGN_JUSTIFIED;
        }
        if (align == CellStyle.VERTICAL_TOP) {
            result = Element.ALIGN_TOP;
        }
        return result;
    }

    protected int getHAlignByExcel(short align) {
        int result = 0;
        if (align == CellStyle.ALIGN_LEFT) {
            result = Element.ALIGN_LEFT;
        }
        if (align == CellStyle.ALIGN_RIGHT) {
            result = Element.ALIGN_RIGHT;
        }
        if (align == CellStyle.ALIGN_JUSTIFY) {
            result = Element.ALIGN_JUSTIFIED;
        }
        if (align == CellStyle.ALIGN_CENTER) {
            result = Element.ALIGN_CENTER;
        }
        return result;
    }
}