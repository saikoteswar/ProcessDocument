package com.example.processdocx.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.converter.core.XWPFConverterException;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import com.lowagie.text.Document;
import com.lowagie.text.pdf.PdfPageEventHelper;
import com.lowagie.text.pdf.PdfWriter;

public class TestPOI {

	public static void main(String[] args) throws XWPFConverterException, FileNotFoundException, IOException, Docx4JException {

		String docxPath = "C:/Sai/BCBSm/XMLTemplatesBCBS/1MEL-Member Escrow Letter ACA_optum.docx";
        String pdfPath = "C:/Sai/BCBSm/BCBSTemplatesOutputJan8/output_7"+String.valueOf(System.currentTimeMillis())+".pdf";
     
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new FileInputStream(docxPath));

        File tempFile = File.createTempFile("wordMLPackage", ".docx");
        tempFile.deleteOnExit();
        wordMLPackage.save(tempFile);

        FileInputStream fileInputStream = new FileInputStream(tempFile);

        XWPFDocument document=new XWPFDocument(fileInputStream);


        for (XWPFParagraph paragraph : document.getParagraphs()) {
        	paragraph.setSpacingBetween(0.8); // Set line spacing to single
//        	paragraph.setSpacingAfterLines(500); // Remove space after paragraph
			paragraph.setSpacingAfter(-60); // Remove space after paragraph
			paragraph.setSpacingBefore(-10); // Remove space before paragraph
//            paragraph.setSpacingLineRule(LineSpacingRule.AUTO); // Set line spacing rule to auto
		}

        PdfOptions options = PdfOptions.create();
        // Set options to retain page layoutl
        options.setConfiguration(pdfWriter -> {
            pdfWriter.setViewerPreferences(PdfWriter.PageModeUseOutlines);
            pdfWriter.setCompressionLevel(9);
            pdfWriter.setFullCompression();
            pdfWriter.setPageEvent(new PdfPageEventHelper() {
                @Override
                public void onEndPage(PdfWriter writer, Document document) {
                    // Ensure content stays on the same page
                    writer.getDirectContent().setLiteral("\n");
                }
            });
        });

        PdfConverter.getInstance().convert(document, new FileOutputStream(pdfPath) , options);
        
}}
