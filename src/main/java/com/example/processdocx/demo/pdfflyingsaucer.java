package com.example.processdocx.demo;

import java.io.File;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class pdfflyingsaucer {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
		String inputDocxPath = "C:/Sai/BCBSm/XMLTemplatesBCBS/new 242.docx";
		String outputPdfPath = "C:/Sai/BCBSm/BCBSTemplatesOutputJan8/output_7"+String.valueOf(System.currentTimeMillis())+".pdf";
		
		   try {
		        // Load the Word document
		        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputDocxPath));

		        // Convert to PDF
		        Docx4J.toPDF(wordMLPackage, null);

		        System.out.println("Conversion completed successfully: " + outputPdfPath);
		    } catch (Exception e) {
		        e.printStackTrace();
		        System.err.println("Error occurred during conversion: " + e.getMessage());
		    }
		}
		


}
