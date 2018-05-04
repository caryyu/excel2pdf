# excel2pdf
An easy way to convert Excel to PDF by Java code based on Apache POI and itextpdf. 利用 JAVA 编写把 Excel 转 PDF 解决方案，依赖POI与IText库的实现。

# 单文件转换 / Single files convert
```
String pathOfXls = "D:\\pdfexport\\MAD 5-3-05-Octavia NF-20131025.xls";
String pathOfPdf = "D:\\pdfexport\\MAD 5-3-05-Octavia NF-20131025.pdf";

FileInputStream fis = new FileInputStream(pathOfXls);
List<ExcelObject> objects = new ArrayList<ExcelObject>();
objects.add(new ExcelObject("导航1",fis));
FileOutputStream fos = new FileOutputStream(pathOfPdf);
Excel2Pdf pdf = new Excel2Pdf(objects, fos);
pdf.convert();
```

# 多文件转换 / Multiple files convert
多文件转换之后会合并至某一个 PDF 中，并且支持导航栏标题方式。
```
FileInputStream fis1 = new FileInputStream(new File("D:\\pdfexport\\MAD 5-3-05-Octavia NF-20131025.xls"));
FileInputStream fis2 = new FileInputStream(new File("D:\\pdfexport\\MAD 6-1-47-Octavia NF-20131025.xls"));
FileInputStream fis3 = new FileInputStream(new File("D:\\pdfexport\\MAD 038-Superb FL DS-20131025.xls"));

FileOutputStream fos = new FileOutputStream(new File("D:\\test.pdf"));

List<ExcelObject> objects = new ArrayList<ExcelObject>();
objects.add(new ExcelObject("1.MAD 5-3-05-Octavia NF-20131025.xls",fis1));
objects.add(new ExcelObject("2.MAD 6-1-47-Octavia NF-20131025.xls",fis2));
objects.add(new ExcelObject("3.MAD 038-Superb FL DS-20131025.xls",fis3));

Excel2Pdf pdf = new Excel2Pdf(objects , fos);
pdf.convert();
```

# 代码打包
```
package -Prelease -Dmaven.test.skip=true
```

# 贡献与建议 / Contribution 
希望拿了此份代码的人改进了一些问题或提交BUG的请提交PR到这个主干上，我会做合并操作希望把这个库进行更加的完善，谢谢！
All Contributions are welcomed

# 后续功能计划（可能需要重构）
1、对单个 Excel 文件的 Sheet 进行 PDF 的页处理，并把 Sheet 名称当做锚；  
2、支持多个 Excel 文件的 Sheet 合并，并在 PDF 页后面进行追加；  
3、实现自动分辨 Excel 版式并对内容进行有效的缩放。  
  