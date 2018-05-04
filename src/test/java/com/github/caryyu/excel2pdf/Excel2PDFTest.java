package com.github.caryyu.excel2pdf;

import com.itextpdf.text.DocumentException;

import java.io.*;
import java.util.*;

import org.junit.*;

public class Excel2PDFTest {
	String resourcesDir = "src/test/resources";
	String outputDir = "target/output";

	@Before
	public void setUp() throws Exception {
		File output = new File(outputDir);
		output.mkdir();
	}

	Map<String, File> getExcelFiles() {
		Map<String, File> excelFiles = new LinkedHashMap<String, File>();
		File dir = new File(resourcesDir);
		if (dir.isDirectory()) {
			File[] files = dir.listFiles(new FilenameFilter() {
				public boolean accept(File dir, String name) {
					return name.toLowerCase().endsWith("xls");
				}
			});
			for (int i = 0; i < files.length; i++) {
				String idx = (i + 1) + ".";
				excelFiles.put(idx + files[i].getName(), files[i]);
			}
		}
		return excelFiles;
	}

	@Test
	public void testCombineAll() throws IOException, DocumentException {
		List<ExcelObject> objects = new ArrayList<ExcelObject>();
		Map<String, File> excelFiles = getExcelFiles();
		for (String index : excelFiles.keySet()) {
			File excelFile = excelFiles.get(index);
			FileInputStream fis = new FileInputStream(excelFile);
			objects.add(new ExcelObject(index, fis));
		}

		FileOutputStream fos = new FileOutputStream(new File(outputDir+"/allInOne.pdf"));
		Excel2Pdf pdf = new Excel2Pdf(objects, fos);

		pdf.convert();
	}

	public static void main(String[] args) throws Exception {
		Excel2PDFTest test = new Excel2PDFTest();
		test.testSingle();
	}

	@Test
	public void testSingle() throws IOException, DocumentException {
		Map<String, File> excelFiles = getExcelFiles();
		for (String index : excelFiles.keySet()) {
			File excelFile = excelFiles.get(index);
			List<ExcelObject> objects = new ArrayList<ExcelObject>();
			FileInputStream fis = new FileInputStream(excelFile);
			FileOutputStream fos = null;
			try {
				objects.add(new ExcelObject(index, fis));
				String excel = excelFile.getName();
				File output = new File(outputDir+"/" + excel.substring(0, excel.lastIndexOf(".")) + ".pdf");
				output.getParentFile().mkdirs();
				fos = new FileOutputStream(output);
				Excel2Pdf pdf = new Excel2Pdf(objects, fos);
				pdf.convert();
			}finally {
				if(fis != null)
					fis.close();
				if(fos != null)
					fos.close();
			}
		}
	}
}