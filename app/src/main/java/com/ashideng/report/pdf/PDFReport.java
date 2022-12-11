package com.ashideng.report.pdf;

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

/**
 * @author zhuhuanhuan@shunjiantech.cn
 * @version 1.0.0
 * @description
 * @create 2022/12/11 下午7:09
 **/
public class PDFReport {

    public String convertToPDF(String path, String pdfPath) {
        try (InputStream in = Files.newInputStream(Paths.get(path));
             XWPFDocument document = new XWPFDocument(in);
             OutputStream outPDF = Files.newOutputStream(Paths.get(pdfPath))) {


            PdfOptions options = PdfOptions.create();
            PdfConverter.getInstance().convert(document, outPDF, options);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return "";
    }
}
