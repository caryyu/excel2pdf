package com.github.caryyu.excel2pdf;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.awt.image.ColorModel;
import java.awt.image.WritableRaster;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * Created by cary on 6/15/17.
 */
public class POIUtil {
	public static int[] getColorRGB(Color color){
        int red = 0;
        int green = 0;
        int blue = 0;

        if (color instanceof HSSFColor) {
            HSSFColor hssfColor = (HSSFColor) color;
            short[] rgb = hssfColor.getTriplet();
            red = rgb[0];
            green = rgb[1];
            blue = rgb[2];
        }else  if (color instanceof XSSFColor) {
            XSSFColor xssfColor = (XSSFColor) color;
            byte[] rgb = xssfColor.getRGB();
            if(rgb != null) {
                red = (rgb[0] < 0) ? (rgb[0] + 256) : rgb[0];
                green = (rgb[1] < 0) ? (rgb[1] + 256) : rgb[1];
                blue = (rgb[2] < 0) ? (rgb[2] + 256) : rgb[2];
            }
        }

        if(red != 0 || green != 0 || blue != 0){
            return new int[] {red,green,blue};
        }else return new int[] {255,255,255};
    }
	
    public static int getRGB(Color color){
        int result = 0x00FFFFFF;

        int red = 0;
        int green = 0;
        int blue = 0;

        if (color instanceof HSSFColor) {
            HSSFColor hssfColor = (HSSFColor) color;
            short[] rgb = hssfColor.getTriplet();
            red = rgb[0];
            green = rgb[1];
            blue = rgb[2];
        }else  if (color instanceof XSSFColor) {
            XSSFColor xssfColor = (XSSFColor) color;
            byte[] rgb = xssfColor.getRGB();
            if(rgb != null) {
                red = (rgb[0] < 0) ? (rgb[0] + 256) : rgb[0];
                green = (rgb[1] < 0) ? (rgb[1] + 256) : rgb[1];
                blue = (rgb[2] < 0) ? (rgb[2] + 256) : rgb[2];
            }
        }

        if(red != 0 || green != 0 || blue != 0){
            result = new java.awt.Color(red, green, blue).getRGB();
        }
        return result;
    }

    public static int getBorderRBG(Workbook wb  , short index){
        int result = 0;

        if(wb instanceof HSSFWorkbook){
            HSSFWorkbook hwb = (HSSFWorkbook)wb;
            HSSFColor color =  hwb.getCustomPalette().getColor(index);
            if(color != null){
                result = getRGB(color);
            }
        }

        if(wb instanceof XSSFWorkbook){
            XSSFColor color = new XSSFColor();
            color.setIndexed(index);
            result = getRGB(color);
        }

        return result;
    }

    @SuppressWarnings("finally")
    public static byte[] scale(byte[] bytes , double width, double height) {
        BufferedImage bufferedImage = null;
        BufferedImage bufTarget = null;
        try {
            ByteArrayInputStream bais = new ByteArrayInputStream(bytes);
            bufferedImage = ImageIO.read(bais);
            double sx =  width / bufferedImage.getWidth();
            double sy =  height / bufferedImage.getHeight();
            int type = bufferedImage.getType();
            if (type == BufferedImage.TYPE_CUSTOM) {
                ColorModel cm = bufferedImage.getColorModel();
                WritableRaster raster = cm.createCompatibleWritableRaster((int)width, (int)height);
                boolean alphaPremultiplied = cm.isAlphaPremultiplied();
                bufTarget = new BufferedImage(cm, raster, alphaPremultiplied, null);
            } else {
                bufTarget = new BufferedImage((int)width, (int)height, type);
            }

            Graphics2D g = bufTarget.createGraphics();
            g.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            g.drawRenderedImage(bufferedImage, AffineTransform.getScaleInstance(sx, sy));
            g.dispose();

            if(bufTarget != null){
                ByteArrayOutputStream baos = new ByteArrayOutputStream();
                ImageIO.write(bufTarget, "png", baos);
                byte[] result = baos.toByteArray();
                return result;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
