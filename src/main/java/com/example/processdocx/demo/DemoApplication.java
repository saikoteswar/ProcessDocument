package com.example.processdocx.demo;

import jakarta.xml.bind.JAXBElement;
import jakarta.xml.bind.JAXBException;
import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.docx4j.jaxb.XPathBinderAssociationIsPartialException;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Text;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@SpringBootApplication
public class DemoApplication implements CommandLineRunner {
	Set<String> sortedSetFreeText = new TreeSet<>();
	//write code to initialize a hashmap
	LinkedHashMap<String, ArrayList<String>> map = new LinkedHashMap<String, ArrayList<String>>();
	ArrayList<String> ftList = new ArrayList<>();


	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {


		if (args[0].equalsIgnoreCase("1")) {
			createReplacementMap();
		} else if (args[0].equalsIgnoreCase("2")) {
			createMergeFieldExcel();
		} else if (args[0].equalsIgnoreCase("3")) {
			createMergeFieldExcelFreeText();
		}
		else if (args[0].equalsIgnoreCase("4")){
			editMergeFieldWordProcessing();
		}
	}

	public void createReplacementMap() throws IOException {

		String outputFilePath;
		// Step 1: Read Excel file into a HashMap
		String excelFilePath = "C:/Sai/BCBSm/BCBSMASS Merge Fields ELG_v3.xlsx";
		Map<String, String> placeholderMap = readExcelFile(excelFilePath);


		// New HashMap with lowercase keys for case-insensitive lookup
		Map<String, String> replacements = new HashMap<>();
		for (Map.Entry<String, String> entry : placeholderMap.entrySet()) {
			String normalizedKey = entry.getKey().toLowerCase().replaceAll("\\s+", "");
			replacements.put(normalizedKey, entry.getValue());
		}

		String inputFolderPath = "C:/Sai/BCBSm/bcbsmass";
		String outputFolderPath = "C:/Sai/BCBSm/BCBSTemplatesOutput1"; // Output folder for processed .docx files

		File inputFolder = new File(inputFolderPath);
		File[] docxFiles = inputFolder.listFiles((dir, name) -> name.endsWith(".docx") && !name.startsWith("~"));


		Workbook workbook = new XSSFWorkbook();


		if (docxFiles != null) {
			for (File docxFile : docxFiles) {


				try {
					String sheetNameToTest = docxFile.getName().length() > 31 ?
							docxFile.getName().substring(0, 31) : docxFile.getName();
					Sheet sheet = doesSheetExist(workbook, sheetNameToTest) ?
							workbook.createSheet(docxFile.getName().substring(0, 28) + "-".concat(RandomStringUtils.randomAlphanumeric(2)))
							: workbook.createSheet(docxFile.getName());

//						Sheet sheet = workbook.createSheet(docxFile.getName().substring(0,30).r);
					Row headerRow = sheet.createRow(1);
					// Create a CellStyle with blue background color
					headerRow.createCell(1).setCellValue("Merge Field");
					headerRow.createCell(2).setCellValue("Replacement Value");

					// Load the DOCX file
					FileInputStream fis = new FileInputStream(docxFile);
					XWPFDocument document = new XWPFDocument(fis);

					// Process each paragraph to find and replace merge fields
					for (XWPFParagraph paragraph : document.getParagraphs()) {
						for (XWPFRun run : paragraph.getRuns()) {
							String text = run.getText(0);
							if (text != null) {
								// Replace merge fields with corresponding values from HashMap
								String modifiedText = replaceMergeFieldsNew(text, replacements, workbook, sheet, headerRow);
								run.setText(modifiedText, 0);
							}
						}
					}

					outputFilePath = outputFolderPath + File.separator + docxFile.getName();
					// Write the updated document to an output file
					FileOutputStream fos = new FileOutputStream(outputFilePath);
					document.write(fos);
					fos.close();
					document.close();
					fis.close();

					System.out.println("Merge fields replaced successfully!");

				} catch (IOException e) {
					e.printStackTrace();
				}

			}
		}

		// Write the Excel file with replacement values
		FileOutputStream fos = new FileOutputStream("C:/Sai/BCBSm/Merge_Fields_Replacements.xlsx");
		workbook.write(fos);
		fos.close();
		workbook.close();

	}

	private Map<String, String> readExcelFile(String filePath) throws IOException {
		Map<String, String> placeholderMap = new HashMap<>();
		try (FileInputStream fis = new FileInputStream(new File(filePath));
			 Workbook workbook = new XSSFWorkbook(fis)) {
			Sheet sheet = workbook.getSheetAt(0);
			for (Row row : sheet) {
				Cell keyCell = row.getCell(1);
				Cell valueCell = row.getCell(2);

				if (keyCell != null && valueCell != null) {
					String key = keyCell.getStringCellValue();
					String value = valueCell.getStringCellValue();
					placeholderMap.put(key, value);
				}
			}
		}
		return placeholderMap;
	}

	// Method to replace merge fields with values from the HashMap
	private static String replaceMergeFieldsNew(String text, Map<String, String> replacements, Workbook workbook, Sheet sheet, Row headerRow) {

		CellStyle borderStyle = workbook.createCellStyle();
		borderStyle.setBorderTop(BorderStyle.THIN);
		borderStyle.setBorderBottom(BorderStyle.THIN);
		borderStyle.setBorderLeft(BorderStyle.THIN);
		borderStyle.setBorderRight(BorderStyle.THIN);


		// Define a header style with a background color and border
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setBorderTop(BorderStyle.THIN);
		headerStyle.setBorderBottom(BorderStyle.THIN);
		headerStyle.setBorderLeft(BorderStyle.THIN);
		headerStyle.setBorderRight(BorderStyle.THIN);

		byte[] rgb = new byte[]{(byte) 218, (byte) 233, (byte) 248}; // Hex color #0066CC
		XSSFColor color = new XSSFColor(rgb, null);
		((XSSFCellStyle) headerStyle).setFillForegroundColor(color);
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		//style to set
		System.out.println(text);
		// Pattern to match merge fields (e.g., MERGEFIELD FirstName \* MERGEFORMAT)
		Pattern pattern = Pattern.compile("«(.*?)»");

		Matcher matcher = pattern.matcher(text);
		StringBuffer modifiedText = new StringBuffer();

		while (matcher.find()) {
			Row row = sheet.createRow(sheet.getLastRowNum() + 1);
			row.createCell(1).setCellValue(matcher.group(1));
			String field = matcher.group(1); // Text between « and »
			String replacementValue = replacements.getOrDefault(field.toLowerCase(), null);

			if (replacementValue != null && !replacementValue.equals("")) {
				// Replace the field with its value if found in map
				if (field.equalsIgnoreCase("City_State_Zip") || field.equalsIgnoreCase("SignatureUserEmailId")) {
					matcher.appendReplacement(modifiedText, Matcher.quoteReplacement(replacementValue));
				} else {
					matcher.appendReplacement(modifiedText, Matcher.quoteReplacement("${" + replacementValue + "}"));
				}
				row.createCell(2).setCellValue(replacementValue);

			} else {
				// Keep the original placeholder if no replacement is found
				matcher.appendReplacement(modifiedText, Matcher.quoteReplacement("«" + field + "»"));
				row.createCell(2).setCellValue("");
			}
			// Apply border style to non-empty cells
			for (Row row1 : sheet) {
				for (Cell cell : row1) {
					if (cell != null && cell.getCellType() != CellType.BLANK) {
						cell.setCellStyle(borderStyle);
					}
				}
			}

			for (Cell cell : headerRow) {
				cell.setCellStyle(headerStyle);
			}
			sheet.autoSizeColumn(1);
			sheet.autoSizeColumn(2);
		}
		matcher.appendTail(modifiedText);
		return modifiedText.toString();
	}

	public static boolean doesSheetExist(Workbook workbook, String sheetName) {
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase(sheetName)) {
				System.out.println("Sheet Name already exists" +sheetName);
				return true;
			}
		}
		return false;
	}

	private void createMergeFieldExcel() throws IOException {

		String inputFolderPath = "C:/Sai/BCBSm/bcbsmass";

		Workbook workbook = new XSSFWorkbook();
		File inputFolder = new File(inputFolderPath);
		File[] docxFiles = inputFolder.listFiles((dir, name) -> name.endsWith(".docx") && !name.startsWith("~"));
		Set<String> sortedSet = new TreeSet<>();
		if (docxFiles != null) {
			for (File docxFile : docxFiles) {

				try {

					FileInputStream fis = new FileInputStream(docxFile);
					XWPFDocument document = new XWPFDocument(fis);
					for (XWPFParagraph paragraph : document.getParagraphs()) {
						for (XWPFRun run : paragraph.getRuns()) {
							String text = run.getText(0);
							if (text != null) {

								Pattern pattern = Pattern.compile("«(.*?)»");
								Matcher matcher = pattern.matcher(text);
								while (matcher.find()) {
									String field = matcher.group(1);
									sortedSet.add(field.replaceAll("\\s+", ""));
									System.out.println(field);
								}
							}
						}
					}
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
		Sheet sheet = workbook.createSheet("TroverisMergeFields");
		Row headerRow = sheet.createRow(1);
		Cell headerCell = headerRow.createCell(1);
		headerCell.setCellValue("Merge Fields");

		int rowIndex = 2;
		for (String value : sortedSet) {
			Row row = sheet.createRow(rowIndex++);
			Cell cell = row.createCell(1);
			cell.setCellValue(value);
		}
		FileOutputStream fos = new FileOutputStream("C:/Sai/BCBSm/MergeFieldsBCBS_1.xlsx");
		workbook.write(fos);
		fos.close();
		workbook.close();
		System.out.println("File printed successfully");
	}

	private void createMergeFieldExcelFreeText() throws IOException {

		String inputFolderPath = "C:/Sai/BCBSm/bcbsmass";

		File inputFolder = new File(inputFolderPath);
		File[] docxFiles = inputFolder.listFiles((dir, name) -> name.endsWith(".docx") && !name.startsWith("~"));

		Workbook workbook = new XSSFWorkbook();

		if (docxFiles != null) {
			for (File docxFile : docxFiles) {
				try {
					// Load the DOCX file
					FileInputStream fis = new FileInputStream(docxFile);
					XWPFDocument document = new XWPFDocument(fis);
					boolean freeTextFound = false;
					System.out.println("Doc Nam is " + docxFile.getName());
					boolean freeTextFoundInLetter = false;
					// Process each paragraph to find and replace merge fields
					for (XWPFParagraph paragraph : document.getParagraphs()) {
						for (XWPFRun run : paragraph.getRuns()) {
							String text = run.getText(0);
							if (text != null) {
								// Replace merge fields with corresponding values from HashMap
								freeTextFound=	replaceMergeFieldsNew_FreeText( text);
								if(freeTextFound){
									freeTextFoundInLetter = true;
								}

							}
						}
					}


//					System.out.println("Merge fields replaced successfully!");
//					System.out.println("Free Text Value "+freeTextFound);
					if(freeTextFoundInLetter){


						CellStyle borderStyle = workbook.createCellStyle();
						borderStyle.setBorderTop(BorderStyle.THIN);
						borderStyle.setBorderBottom(BorderStyle.THIN);
						borderStyle.setBorderLeft(BorderStyle.THIN);
						borderStyle.setBorderRight(BorderStyle.THIN);


						// Define a header style with a background color and border
						CellStyle headerStyle = workbook.createCellStyle();
						headerStyle.setBorderTop(BorderStyle.THIN);
						headerStyle.setBorderBottom(BorderStyle.THIN);
						headerStyle.setBorderLeft(BorderStyle.THIN);
						headerStyle.setBorderRight(BorderStyle.THIN);

						byte[] rgb = new byte[]{(byte) 218, (byte) 233, (byte) 248}; // Hex color #0066CC
						XSSFColor color = new XSSFColor(rgb, null);
						((XSSFCellStyle) headerStyle).setFillForegroundColor(color);
						headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

						String sheetNameToTest = docxFile.getName().length() > 31 ? docxFile.getName().substring(0, 31) : docxFile.getName();
						System.out.println("Sheet Name "+sheetNameToTest);
						boolean docExists = false;
						docExists = doesSheetExist(workbook, sheetNameToTest);
						Sheet sheet = null;
						if(docExists){
							sheet=	workbook.createSheet(docxFile.getName().substring(0, 28) + "-"
									.concat(RandomStringUtils.randomAlphanumeric(2)));
						}else{
							sheet = 	workbook.createSheet(docxFile.getName());
						}

						Row headerRow = sheet.createRow(1);
						// Create a CellStyle with blue background color
						headerRow.createCell(1).setCellValue("Merge Field");


						int rowIndex = 2;
						ftList = new ArrayList<>();
						for (String value : sortedSetFreeText) {
							Row row = sheet.createRow(rowIndex++);
							Cell cell = row.createCell(1);
							cell.setCellValue(value);
							ftList.add(value);
						}
						map.put(docxFile.getName(),ftList);

						sortedSetFreeText.clear();

						// Apply border style to non-empty cells
						for (Row row1 : sheet) {
							for (Cell cell : row1) {
								if (cell != null && cell.getCellType() != CellType.BLANK) {
									cell.setCellStyle(borderStyle);
								}
							}
						}

						for (Cell cell : headerRow) {
							cell.setCellStyle(headerStyle);
						}
						sheet.autoSizeColumn(1);
						sheet.autoSizeColumn(2);

					}

				} catch (IOException e) {
					e.printStackTrace();
				}

			}
		}

		System.out.println("Process completed");
		// Write the Excel file with replacement values
		FileOutputStream fos = new FileOutputStream("C:/Sai/BCBSm/Merge_Fields_Replacements_FreeText.xlsx");
		workbook.write(fos);
		fos.close();
		workbook.close();

//		Write data from Hashmap into excel
		Workbook workbook1 = new XSSFWorkbook();

		Sheet sheet = workbook1.createSheet("FreeTextValues");

		int rowno=0;

		for(HashMap.Entry entry:map.entrySet()) {
			Row row=sheet.createRow(rowno++);
			row.createCell(0).setCellValue((String)entry.getKey());

			for(int i=0;i<((ArrayList<String>)entry.getValue()).size();i++){
				row.createCell(i+1).setCellValue( ((ArrayList<String>)entry.getValue()).get(i));
			}
		}

		FileOutputStream fos1 = new FileOutputStream("C:/Sai/BCBSm/Merge_Fields_Replacements_FreeText_consolicated.xlsx");
		workbook1.write(fos1);
		fos1.close();
		workbook1.close();
	}

	private boolean  replaceMergeFieldsNew_FreeText( String text ) {


		boolean freeTextFound = false;

		//style to set
//		System.out.println(text);
		// Pattern to match merge fields (e.g., MERGEFIELD FirstName \* MERGEFORMAT)

		Pattern pattern = Pattern.compile("«(.*?)»");

		Matcher matcher = pattern.matcher(text);

		while (matcher.find()) {
			String field = matcher.group(1); // Text between « and »
			System.out.println("Field is "+field);
			if (field.startsWith("FT_")) {
				sortedSetFreeText.add(field);
				freeTextFound = true;
			}
		}

		return freeTextFound;
	}

	private void editMergeFieldWordProcessing() throws Docx4JException, JAXBException {

		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File("C:\\Sai\\BCBSm\\BCBSTemplates/1MEL-Member Escrow Letter ACA.docx"));

		// Get the document's main content (paragraphs, tables, etc.)
		MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();

		// Create a map with merge field keys and values to replace
		Map<String, String> replacements = Map.of(
				"«Date_of_letter»", "${DATE_SENT}"
		);

		// Replace merge fields with corresponding values
		replaceMergeFields(documentPart, replacements);

		// Save the modified document
		wordMLPackage.save(new File("C:\\Sai\\BCBSm\\BCBSTemplates/lmodified_document.docx"));

	}
	
	public static void replaceMergeFields(MainDocumentPart documentPart, Map<String, String> replacements) throws JAXBException, XPathBinderAssociationIsPartialException {

		List<Object> texts = documentPart.getJAXBNodesViaXPath("//w:t", true);


		// Find and replace all merge fields (e.g., <<FieldName>>)
		for (Object obj : texts) {
			if (obj instanceof JAXBElement) {
				JAXBElement<?> jaxbElement = (JAXBElement<?>) obj;

				// Check for fldSimple elements (merge fields in Word)
				if (jaxbElement.getValue() instanceof Text) {
					Text textElement = (Text) jaxbElement.getValue();
					String textContent = textElement.getValue();
					System.out.println(textContent);
					for (String key : replacements.keySet()) {
						if (textContent.contains(key)) {
							textContent = textContent.replace(key, replacements.get(key));
							textElement.setValue(textContent);
						}
					}
				}
			}
		}
	}

}
