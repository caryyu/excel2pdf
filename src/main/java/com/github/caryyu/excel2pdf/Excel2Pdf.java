package com.github.caryyu.excel2pdf;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;

import java.io.IOException;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by cary on 6/15/17.
 */
public class Excel2Pdf extends PdfTool{
    protected List<ExcelObject> objects = new ArrayList<ExcelObject>();

    /**
     * <p>Description: 导出单项PDF，不包含目录</p>
     * @param object
     */
    public Excel2Pdf(ExcelObject object , OutputStream os) {
        this.objects.add(object);
        this.os = os;
    }

    /**
     * <p>Description: 导出多项PDF，包含目录</p>
     * @param objects
     */
    public Excel2Pdf(List<ExcelObject> objects , OutputStream os) {
        this.objects = objects;
        this.os = os;
    }

    /**
     * <p>Description: 转换调用</p>
     * @throws DocumentException
     * @throws MalformedURLException
     * @throws IOException
     */
    public void convert() throws DocumentException, MalformedURLException, IOException {
        getDocument().setPageSize(PageSize.A4.rotate());
        PdfWriter writer = PdfWriter.getInstance(getDocument(), os);
        writer.setPageEvent(new PDFPageEvent());
        //Open document
        getDocument().open();
        //Single one
        if(this.objects.size() <= 1){
            PdfPTable table = this.toCreatePdfTable(this.objects.get(0) ,  getDocument() , writer);
            getDocument().add(table);
        }
        //Multiple ones
        if(this.objects.size() > 1){
            toCreateContentIndexes(writer , this.getDocument() , this.objects);
            //
            for (int i = 0; i < this.objects.size(); i++) {
                PdfPTable table = this.toCreatePdfTable(this.objects.get(i) , getDocument() , writer);
                getDocument().add(table);
            }
        }
        //
        getDocument().close();
    }

    protected PdfPTable toCreatePdfTable(ExcelObject object , Document document , PdfWriter writer) throws MalformedURLException, IOException, DocumentException{
        PdfPTable table = new PdfTableExcel(object).getTable();
        table.setKeepTogether(true);
//      table.setWidthPercentage(new float[]{100} , writer.getPageSize());
        table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
        return table;
    }

    /**
     * <p>Description: 内容索引创建</p>
     * @throws DocumentException
     */
    protected void toCreateContentIndexes(PdfWriter writer , Document document , List<ExcelObject> objects) throws DocumentException{
        PdfPTable table = new PdfPTable(1);
        table.setKeepTogether(true);
        table.getDefaultCell().setBorder(PdfPCell.NO_BORDER);
        //
        Font font = new Font(Resource.BASE_FONT_CHINESE , 12 , Font.NORMAL);
        font.setColor(new BaseColor(0,0,255));
        //
        for (int i = 0; i < objects.size(); i++) {
            ExcelObject o = objects.get(i);
            String text = o.getAnchorName();
            Anchor anchor = new Anchor(text , font);
            anchor.setReference("#" + o.getAnchorName());
            //
            PdfPCell cell = new PdfPCell(anchor);
            cell.setBorder(0);
            //
            table.addCell(cell);
        }
        //
        document.add(table);
    }

    /**
     * <p>ClassName: PDFPageEvent</p>
     * <p>Description: 事件 -> 页码控制</p>
     * <p>Author: Cary</p>
     * <p>Date: Oct 25, 2013</p>
     */
    private static class PDFPageEvent extends PdfPageEventHelper {
        protected PdfTemplate template;
        public BaseFont baseFont;

        @Override
        public void onStartPage(PdfWriter writer, Document document) {
            try{
                this.template = writer.getDirectContent().createTemplate(100, 100);
                this.baseFont = new Font(Resource.BASE_FONT_CHINESE , 8, Font.NORMAL).getBaseFont();
            } catch(Exception e) {
                throw new ExceptionConverter(e);
            }
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            //在每页结束的时候把“第x页”信息写道模版指定位置
            PdfContentByte byteContent = writer.getDirectContent();
            String text = "第" + writer.getPageNumber() + "页";
            float textWidth = this.baseFont.getWidthPoint(text, 8);
            float realWidth = document.right() - textWidth;
            //
            byteContent.beginText();
            byteContent.setFontAndSize(this.baseFont , 10);
            byteContent.setTextMatrix(realWidth , document.bottom());
            byteContent.showText(text);
            byteContent.endText();
            byteContent.addTemplate(this.template , realWidth , document.bottom());
        }
    }
}