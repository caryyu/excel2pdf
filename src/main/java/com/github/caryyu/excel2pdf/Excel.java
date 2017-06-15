package com.github.caryyu.excel2pdf;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

/**
 * Created by cary on 6/15/17.
 */
public class Excel {

    protected Workbook wb;
    protected Sheet sheet;

    public Excel(InputStream is) {
        try {
            this.wb = WorkbookFactory.create(is);
            this.sheet = wb.getSheetAt(wb.getActiveSheetIndex());
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    public Sheet getSheet() {
        return sheet;
    }

    public Workbook getWorkbook(){
        return wb;
    }
}
