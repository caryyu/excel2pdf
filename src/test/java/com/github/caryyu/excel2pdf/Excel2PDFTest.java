package com.github.caryyu.excel2pdf;

import com.itextpdf.text.DocumentException;
import org.junit.Test;

import java.io.*;
import java.util.*;

public class Excel2PDFTest {
	String testExcelDir = "testexcel";
//	String[][] excelFiles = { 
//			{ "1.薪資明細表", "testexcel/salary.xls" },
//			{ "2.工時明細表", "testexcel/salarytimerecords.xls" },
//			{ "3.加班明細表", "testexcel/overtimerecords.xls" },
//			{ "3.悠遊卡樣本", "testexcel/悠遊卡樣本.xls" }
//	};

	public void setup() {

	}

	Map<String,File> getExcelFiles(){
		Map<String,File> excelFiles = new LinkedHashMap<String,File>();
		File dir = new File(testExcelDir);
		if(dir.isDirectory()) {
			File[] files = dir.listFiles(new FilenameFilter() {
				public boolean accept(File dir, String name) {
					return name.toLowerCase().endsWith("xls");
				}
			});
			for(int i=0;i<files.length;i++) {
				String idx = (i+1)+".";
				excelFiles.put(idx+files[i].getName(),files[i]);
			}
		}
		return excelFiles;
	}
	@Test
	public void testCombineAll() throws IOException, DocumentException {
		List<ExcelObject> objects = new ArrayList<ExcelObject>();
		Map<String,File> excelFiles = getExcelFiles();
		for(String index:excelFiles.keySet()) {
			File excelFile = excelFiles.get(index);
			FileInputStream fis = new FileInputStream(excelFile);
			objects.add(new ExcelObject(index, fis));
		}

		FileOutputStream fos = new FileOutputStream(new File("testpdf/allInOne.pdf"));
		Excel2Pdf pdf = new Excel2Pdf(objects, fos);

		pdf.convert();
	}

	public static void main(String[] args) throws Exception {
		Excel2PDFTest test = new Excel2PDFTest();
		test.testSingle();
	}

	@Test
	public void testSingle() throws IOException, DocumentException {
		Map<String,File> excelFiles = getExcelFiles();
		for(String index:excelFiles.keySet()) {
			File excelFile = excelFiles.get(index);
			List<ExcelObject> objects = new ArrayList<ExcelObject>();
			FileInputStream fis = new FileInputStream(excelFile);
			objects.add(new ExcelObject(index, fis));
			String excel = excelFile.getName();
			File output = new File("testpdf/" + excel.substring(0, excel.lastIndexOf(".")) + ".pdf");
			output.getParentFile().mkdirs();
			FileOutputStream fos = new FileOutputStream(output);
			Excel2Pdf pdf = new Excel2Pdf(objects, fos);
			pdf.convert();
		}
	}
}