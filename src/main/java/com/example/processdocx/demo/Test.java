package com.example.processdocx.demo;



import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.docx4j.Docx4J;
import org.docx4j.convert.out.FOSettings;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import jakarta.xml.bind.JAXBException;

public class Test {

	public static void main(String[] args) throws JAXBException, Docx4JException, IOException {

//        String docxInputPath = "C:/Sai/BCBSm/IndentationTesting/new 237.docx";
        String docxInputPath = "C:/Sai/BCBSm/XMLTemplatesBCBS/new 242.docx";
        String pdfPath = "C:/Sai/BCBSm/BCBSTemplatesOutputJan8/output_7"+String.valueOf(System.currentTimeMillis())+".pdf";
     
        convertDocxToPdf(docxInputPath, pdfPath);
	}

	private static void convertDocxToPdf(String docxInputPath, String pdfPath) {

		try {
			// Load the Word document
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(docxInputPath));

			// Create FOSettings instance
			FOSettings foSettings = Docx4J.createFOSettings();

			// Set up FOSettings properties
			foSettings.setWmlPackage(wordMLPackage); // Associate the Word package
			foSettings.setApacheFopMime("application/pdf"); // Set output type to PDF

			// Optional: Disable debug output
			foSettings.setOpcPackage(wordMLPackage);

			// Convert to PDF
			Docx4J.toFO(foSettings, new FileOutputStream("C:/Sai/BCBSm/BCBSTemplatesOutputJan8/output_7"+String.valueOf(System.currentTimeMillis())+".pdf"), 
					Docx4J.FLAG_EXPORT_PREFER_XSL);

			System.out.println("PDF conversion completed successfully.");

		} catch (Exception e) {
			e.printStackTrace();
		}

	}
}
