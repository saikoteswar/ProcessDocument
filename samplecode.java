spring:
  kafka:
    bootstrap-servers: your-gcp-kafka-broker-1:9093,your-gcp-kafka-broker-2:9093
    producer:
      key-serializer: org.apache.kafka.common.serialization.StringSerializer
      value-serializer: org.apache.kafka.common.serialization.StringSerializer
    consumer:
      group-id: your-consumer-group
      key-deserializer: org.apache.kafka.common.serialization.StringDeserializer
      value-deserializer: org.apache.kafka.common.serialization.StringDeserializer
    properties:
      security.protocol: SSL
      ssl:
        # Truststore (CA certificate)
        truststore:
          location: classpath:truststore.jks
          password: ${TRUSTSTORE_PASSWORD}  # Use env variable
        # Keystore (client certificate + private key)
        keystore:
          location: classpath:keystore.jks
          password: ${KEYSTORE_PASSWORD}    # Use env variable
        # Private key password (if different from keystore password)
        key:
          password: ${KEY_PASSWORD}
        # Disable hostname verification (for testing only)
        endpoint:
          identification:
            algorithm: ""


--------

# Kafka Bootstrap Server
spring.kafka.bootstrap-servers=your-gcp-kafka-server:9093

# SSL Configuration
spring.kafka.properties.security.protocol=SSL
spring.kafka.properties.ssl.truststore.location=classpath:truststore.jks
spring.kafka.properties.ssl.truststore.password=truststore-password
spring.kafka.properties.ssl.keystore.location=classpath:keystore.jks
spring.kafka.properties.ssl.keystore.password=keystore-password
spring.kafka.properties.ssl.key.password=key-password  # If different from keystore password

# Disable hostname verification (if needed)
spring.kafka.properties.ssl.endpoint.identification.algorithm=



				 @Configuration
public class KafkaConfig {

    @Value("${spring.kafka.bootstrap-servers}")
    private String bootstrapServers;

    @Bean
    public ProducerFactory<String, String> producerFactory() {
        Map<String, Object> configProps = new HashMap<>();
        configProps.put(ProducerConfig.BOOTSTRAP_SERVERS_CONFIG, bootstrapServers);
        configProps.put(ProducerConfig.KEY_SERIALIZER_CLASS_CONFIG, StringSerializer.class);
        configProps.put(ProducerConfig.VALUE_SERIALIZER_CLASS_CONFIG, StringSerializer.class);
        return new DefaultKafkaProducerFactory<>(configProps);
    }

    @Bean
    public KafkaTemplate<String, String> kafkaTemplate() {
        return new KafkaTemplate<>(producerFactory());
    }
}

----------

	


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*
	;

public class ExcelReader {
    public static void main(String[] args) {
        String filePath = "mergeFieldsreplacement.xlsx"; // Path to your Excel file

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet
            Map<String, List<String>> keyValueMap = new HashMap<>();

            for (Row row : sheet) {
                if (row.getRowNum() == 0) continue; // Skip header row

                Cell keyCell = row.getCell(1); // Column B is the key
                if (keyCell == null || keyCell.getCellType() == CellType.BLANK) continue;

                String key = keyCell.toString().trim();
                List<String> values = new ArrayList<>();

                for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                    if (colIndex == 1) continue; // Skip key column (Column B)
                    Cell cell = row.getCell(colIndex);
                    if (cell != null) {
                        values.add(cell.toString().trim());
                    }
                }

                keyValueMap.put(key, values);
            }

            // Iterate over the HashMap and print key-value pairs in the required format
            for (Map.Entry<String, List<String>> entry : keyValueMap.entrySet()) {
                String key = entry.getKey();
                List<String> values = entry.getValue();

                for (String value : values) {
                    System.out.println(key + ", " + value);
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}

-----------------
	
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class FileChecker {
    public static void main(String[] args) {
        String excelFilePath = "C:\\path\\to\\your\\file.xlsx"; // Update with your Excel file path
        String folderPath = "C:\\path\\to\\your\\folder"; // Update with your folder path
        int columnIndex = 0; // Update with the correct column index (0-based index)

        try {
            // Read file names from Excel (removing extra spaces)
            Set<String> excelFileNames = readExcelFileNames(excelFilePath, columnIndex);

            // Read file names from local folder
            Set<String> folderFileNames = readLocalFolderFiles(folderPath);

            // Find files that are in the folder but not in the Excel
            folderFileNames.removeAll(excelFileNames);

            // Print missing files
            if (folderFileNames.isEmpty()) {
                System.out.println("All files in the folder are listed in the Excel file.");
            } else {
                System.out.println("Files in the folder but not in Excel:");
                for (String file : folderFileNames) {
                    System.out.println(file);
                }
            }
        } catch (IOException e) {
            System.err.println("Error reading files: " + e.getMessage());
        }
    }

    // Method to read file names from an Excel column (removing spaces)
    private static Set<String> readExcelFileNames(String filePath, int columnIndex) throws IOException {
        Set<String> fileNames = new HashSet<>();
        FileInputStream fis = new FileInputStream(new File(filePath));
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0); // Read the first sheet

        for (Row row : sheet) {
            Cell cell = row.getCell(columnIndex); // Get the column with file names
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String fileName = cell.getStringCellValue().trim(); // Remove leading & trailing spaces
                if (!fileName.isEmpty()) { // Ignore empty cells
                    fileNames.add(fileName);
                }
            }
        }

        workbook.close();
        fis.close();
        return fileNames;
    }

    // Method to read file names from a local folder
    private static Set<String> readLocalFolderFiles(String folderPath) {
        Set<String> fileNames = new HashSet<>();
        File folder = new File(folderPath);
        File[] files = folder.listFiles();

        if (files != null) {
            for (File file : files) {
                if (file.isFile()) {
                    fileNames.add(file.getName().trim()); // Trim spaces from filenames
                }
            }
        }
        return fileNames;
    }
}

--------------
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.P;
import java.io.File;
import java.util.List;

public class RemoveEmptyParagraph {
    
    public static void main(String[] args) {
        String inputFilePath = "input.docx";
        String outputFilePath = "output.docx";
        String searchText = "Find this text";  // Text to search for
        
        try {
            // Load the .docx file
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputFilePath));
            List<Object> paragraphs = wordMLPackage.getMainDocumentPart().getContent();

            // Find the index of the paragraph containing the search text
            int foundIndex = -1;
            for (int i = 0; i < paragraphs.size(); i++) {
                Object obj = paragraphs.get(i);
                if (obj instanceof P) {
                    P paragraph = (P) obj;
                    String paraText = paragraph.toString();

                    if (paraText.contains(searchText)) {
                        foundIndex = i;
                        break;  // Stop after finding the first occurrence
                    }
                }
            }

            // If found, check the next paragraph and remove it if empty
            if (foundIndex != -1 && foundIndex + 1 < paragraphs.size()) {
                Object nextObj = paragraphs.get(foundIndex + 1);
                if (nextObj instanceof P) {
                    P nextPara = (P) nextObj;
                    if (nextPara.toString().trim().isEmpty()) {
                        paragraphs.remove(foundIndex + 1);
                        System.out.println("Empty paragraph removed.");
                    }
                }
            }

            // Save the modified document
            wordMLPackage.save(new File(outputFilePath));
            System.out.println("Updated document saved as: " + outputFilePath);

        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }
}

--------------
import com.lowagie.text.Document;
import com.lowagie.text.pdf.PdfPageEventHelper;
import com.lowagie.text.pdf.PdfWriter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class TestPOI {

    public static void main(String[] args) throws Exception {
        String docName = "NewFax UCSBA-Updated CSB Attorney.docx";
        String docxPath = "C:/Sai/BCBSm/XMLTemplatesBCBS/Feb9/" + docName;
        String pdfPath = "C:/Sai/BCBSm/BCBSTemplatesOutputJan8/output" + docName + System.currentTimeMillis() + ".pdf";

        // Load the document
        FileInputStream fileInputStream = new FileInputStream(docxPath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        // Handle hyperlinks in footers
        cleanHyperlinksInFooters(document);

        // Configure PDF options
        PdfOptions options = PdfOptions.create();
        options.setConfiguration(pdfWriter -> {
            pdfWriter.setViewerPreferences(PdfWriter.PageModeUseOutlines);
            pdfWriter.setCompressionLevel(9);
            pdfWriter.setFullCompression();
            pdfWriter.setPageEvent(new PdfPageEventHelper() {
                @Override
                public void onEndPage(PdfWriter writer, Document document) {
                    writer.getDirectContent().setLiteral("\n");
                }
            });
        });

        // Convert the document to PDF
        PdfConverter.getInstance().convert(document, new FileOutputStream(pdfPath), options);
        System.out.println("Document converted successfully: " + pdfPath);
    }

    /**
     * Cleans hyperlinks from footers while retaining text and formatting.
     */
    private static void cleanHyperlinksInFooters(XWPFDocument document) {
        List<XWPFFooter> footers = document.getFooterList();

        for (XWPFFooter footer : footers) {
            for (XWPFParagraph paragraph : footer.getParagraphs()) {
                cleanHyperlinksFromParagraph(paragraph);
            }
        }
    }

    /**
     * Removes hyperlink runs from a paragraph and retains text and formatting.
     */
    private static void cleanHyperlinksFromParagraph(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();

        if (runs != null) {
            StringBuilder textBuffer = new StringBuilder();
            for (XWPFRun run : runs) {
                textBuffer.append(run.text());
            }

            // Clear the paragraph and reinsert text without hyperlinks
            paragraph.getRuns().clear();

            if (textBuffer.length() > 0) {
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(textBuffer.toString());

                // Optionally copy formatting properties (can be customized)
                if (!runs.isEmpty()) {
                    copyFormatting(runs.get(0), newRun);
                }
            }
        }
    }

    /**
     * Copies formatting properties from one run to another.
     */
    private static void copyFormatting(XWPFRun sourceRun, XWPFRun targetRun) {
        if (sourceRun.getFontSize() != -1) {
            targetRun.setFontSize(sourceRun.getFontSize());
        }
        if (sourceRun.getFontFamily() != null) {
            targetRun.setFontFamily(sourceRun.getFontFamily());
        }
        targetRun.setBold(sourceRun.isBold());
        targetRun.setItalic(sourceRun.isItalic());
    }
}


______
if (runs != null) {
            StringBuilder textBuffer = new StringBuilder();
            for (XWPFRun run : runs) {
                textBuffer.append(run.text());
            }

            // Clear the paragraph and reinsert text without hyperlinks
            paragraph.getRuns().clear();

            if (textBuffer.length() > 0) {
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(textBuffer.toString());

                // Optionally copy formatting properties (can be customized)
                if (!runs.isEmpty()) {
                    copyFormatting(runs.get(0), newRun);
                }
            }
__________
/**
     * Copies formatting properties from one run to another.
     */
    private static void copyFormatting(XWPFRun sourceRun, XWPFRun targetRun) {
        if (sourceRun.getFontSize() != -1) {
            targetRun.setFontSize(sourceRun.getFontSize());
        }
        if (sourceRun.getFontFamily() != null) {
            targetRun.setFontFamily(sourceRun.getFontFamily());
        }
        targetRun.setBold(sourceRun.isBold());
        targetRun.setItalic(sourceRun.isItalic());
    }
___________
if (runs != null) {
                    for (int i = 0; i < runs.size(); i++) {
                        XWPFRun run = runs.get(i);
                        // Check if the run is inside a hyperlink element
                        if (run.getParent() instanceof XWPFHyperlinkRun) {
                            XWPFHyperlinkRun hyperlinkRun = (XWPFHyperlinkRun) run;

                            // Extract text and remove hyperlink
                            String text = hyperlinkRun.text();
                            XWPFRun newRun = paragraph.insertNewRun(i);
                            newRun.setText(text);

                            // Retain formatting properties
                            copyFormatting(run, newRun);

                            paragraph.removeRun(i + 1); // Remove the original hyperlink run
                        }
                    }
                }
------__
import com.lowagie.text.BadElementException;
import com.lowagie.text.Document;
import com.lowagie.text.pdf.PdfPageEventHelper;
import com.lowagie.text.pdf.PdfWriter;
import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;

public class TestPOI {

    public static void main(String[] args) throws Exception {
        String docName = "NewFax UCSBA-Updated CSB Attorney.docx";
        String docxPath = "C:/Sai/BCBSm/XMLTemplatesBCBS/Feb9/" + docName;
        String pdfPath = "C:/Sai/BCBSm/BCBSTemplatesOutputJan8/output" + docName + System.currentTimeMillis() + ".pdf";

        // Load the document
        FileInputStream fileInputStream = new FileInputStream(docxPath);
        XWPFDocument document = new XWPFDocument(fileInputStream);

        // Handle hyperlinks in footers
        removeHyperlinksInFooter(document);

        // Adjust paragraph formatting if needed
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            paragraph.setSpacingBetween(1.0); // Set single line spacing
        }

        // Configure PDF options
        PdfOptions options = PdfOptions.create();
        options.setConfiguration(pdfWriter -> {
            pdfWriter.setViewerPreferences(PdfWriter.PageModeUseOutlines);
            pdfWriter.setCompressionLevel(9);
            pdfWriter.setFullCompression();
            pdfWriter.setPageEvent(new PdfPageEventHelper() {
                @Override
                public void onEndPage(PdfWriter writer, Document document) {
                    writer.getDirectContent().setLiteral("\n");
                }
            });
        });

        // Convert the document to PDF
        PdfConverter.getInstance().convert(document, new FileOutputStream(pdfPath), options);
        System.out.println("Document converted successfully: " + pdfPath);
    }

    /**
     * Removes hyperlinks from footers while retaining their text and formatting.
     */
    private static void removeHyperlinksInFooter(XWPFDocument document) {
        List<XWPFFooter> footers = document.getFooterList();
        for (XWPFFooter footer : footers) {
            for (XWPFParagraph paragraph : footer.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                if (runs != null) {
                    for (int i = 0; i < runs.size(); i++) {
                        XWPFRun run = runs.get(i);
                        if (run.getCTR().getHyperlinkId() != null) {
                            // Remove the hyperlink by copying text to a new run
                            String text = run.text();
                            XWPFRun newRun = paragraph.insertNewRun(i);
                            newRun.setText(text);

                            // Copy formatting from the original run
                            if (run.getFontSize() != -1) newRun.setFontSize(run.getFontSize());
                            if (run.getFontFamily() != null) newRun.setFontFamily(run.getFontFamily());
                            newRun.setBold(run.isBold());
                            newRun.setItalic(run.isItalic());

                            paragraph.removeRun(i + 1); // Remove the original hyperlink run
                        }
                    }
                }
            }
        }
    }
}

------------------
import com.lowagie.text.BadElementException;
import com.lowagie.text.Document;
import com.lowagie.text.Image;
import com.lowagie.text.pdf.PdfPageEventHelper;
import com.lowagie.text.pdf.PdfWriter;
import fr.opensagres.poi.xwpf.converter.core.XWPFConverterException;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.io.*;
import java.util.List;



public class TestPOI {

    public static void main(String[] args) throws XWPFConverterException, FileNotFoundException, IOException, Docx4JException, InvalidFormatException, BadElementException, BadElementException, InvalidFormatException, BadElementException, InvalidFormatException {
        String docName = "NewFax UCSBA-Updated CSB Attorney.docx";
        String docxPath = "C:/Sai/BCBSm/XMLTemplatesBCBS/Feb9/"+docName;
        String pdfPath = "C:/Sai/BCBSm/BCBSTemplatesOutputJan8/output"+docName + String.valueOf(System.currentTimeMillis()) + ".pdf";

        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new FileInputStream(docxPath));

        File tempFile = File.createTempFile("wordMLPackage", ".docx");
        tempFile.deleteOnExit();
        wordMLPackage.save(tempFile);

        FileInputStream fileInputStream = new FileInputStream(tempFile);

        XWPFDocument document=new XWPFDocument(fileInputStream);

        document.removeBodyElement(0);
        for (XWPFParagraph paragraph : document.getParagraphs()) {
//            paragraph.setFirstLineIndent(-500);
//            paragraph.setSpacingAfter(0); // Remove space after paragraph
//            paragraph.setSpacingBefore(-200); // Remove space before paragraph
//            paragraph.setBorderTop(Borders.ZIG_ZAG_STITCH);
            paragraph.setSpacingBetween(1.0); // Set line spacing to single
//        	paragraph.setSpacingAfterLines(500); // Remove space after paragraph
//			paragraph.setSpacingAfter(-40); // Remove space after paragraph
//			paragraph.setSpacingBefore(-10); // Remove space before paragraph
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
/*

        List<XWPFPictureData> pictures = document.getAllPictures();
        for (XWPFPictureData picture : pictures) {
            byte[] bytes = picture.getData();
            Image img = Image.getInstance(bytes);
            img.setAlignment(Image.ALIGN_RIGHT); // Set alignment to center
            img.setAbsolutePosition(100, 500); // Set the position of the image (adjust as needed)
            document.addPictureData(bytes, XWPFDocument.PICTURE_TYPE_PNG);
        }
*/


/*


        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (XWPFRun run : runs) {
                    List<XWPFPicture> pictures = run.getEmbeddedPictures();
                    for (XWPFPicture picture : pictures) {
                        XWPFPictureData pictureData = picture.getPictureData();
                        byte[] bytes = pictureData.getData();
                        int pictureType = pictureData.getPictureType();
                        String filename = pictureData.getFileName();
                        run.addPicture(new ByteArrayInputStream(bytes), pictureType, filename, Units.toEMU(100), Units.toEMU(100)); // Adjust width and height as needed
                    }
                }
            }
        }

*/

        PdfConverter.getInstance().convert(document, new FileOutputStream(pdfPath) , options);
        System.out.println("Document converted succesfully " + pdfPath);
    }
}


----------------------------
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.List;

public class RemoveHyperlinksWithFooterCheck {

    public static void main(String[] args) {
        String inputPath = "path/to/input.docx";
        String outputPath = "path/to/output.docx";

        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputPath));

            // Ensure document has footers, and remove hyperlinks
            ensureFooterExistsAndRemoveHyperlinks(wordMLPackage);

            // Save the updated document
            wordMLPackage.save(new File(outputPath));

            System.out.println("Hyperlinks removed successfully and saved to: " + outputPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void ensureFooterExistsAndRemoveHyperlinks(WordprocessingMLPackage wordMLPackage) throws Docx4JException {
        List<SectionWrapper> sections = wordMLPackage.getDocumentModel().getSections();

        for (SectionWrapper sectionWrapper : sections) {
            FooterPart footerPart = sectionWrapper.getFooter();

            if (footerPart == null) {
                // Create a new footer if one doesn't exist
                footerPart = new FooterPart();
                footerPart.setPackage(wordMLPackage);
                wordMLPackage.getMainDocumentPart()
                             .addTargetPart(footerPart);

                Footer footer = new ObjectFactory().createFooter();
                P paragraph = new ObjectFactory().createP();
                footer.getContent().add(paragraph);
                footerPart.setJaxbElement(footer);

                sectionWrapper.getSectPr()
                              .getEGHdrFtrReferences()
                              .add(footerPart.getRelLast().getReference());
            }

            removeHyperlinksFromFooterContent(footerPart.getContent());
        }
    }

    private static void removeHyperlinksFromFooterContent(List<Object> footerContent) {
        for (Object obj : footerContent) {
            if (obj instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) obj;
                Object value = element.getValue();

                if (value instanceof P) {
                    removeHyperlinksFromParagraph((P) value);
                }
            }
        }
    }

    private static void removeHyperlinksFromParagraph(P paragraph) {
        List<Object> children = paragraph.getContent();

        for (int i = 0; i < children.size(); i++) {
            Object child = children.get(i);
            if (child instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) child;
                Object value = element.getValue();

                if (value instanceof P.Hyperlink) {
                    P.Hyperlink hyperlink = (P.Hyperlink) value;

                    // Extract and replace hyperlink text
                    R formattedTextRun = extractTextWithFormatting(hyperlink.getContent());
                    paragraph.getContent().set(i, formattedTextRun);
                }
            }
        }
    }

    private static R extractTextWithFormatting(List<Object> content) {
        ObjectFactory factory = new ObjectFactory();
        R newRun = factory.createR();

        for (Object child : content) {
            if (child instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) child;
                if (element.getValue() instanceof R) {
                    R originalRun = (R) element.getValue();

                    // Copy run properties (font, size, bold, etc.)
                    if (originalRun.getRPr() != null) {
                        newRun.setRPr(originalRun.getRPr());
                    }

                    for (Object runContent : originalRun.getContent()) {
                        if (runContent instanceof Text) {
                            Text textElement = factory.createText();
                            textElement.setValue(((Text) runContent).getValue());
                            newRun.getContent().add(textElement);
                        }
                    }
                }
            }
        }
        return newRun;
    }
}

-------------------
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.List;

public class RemoveHyperlinksInFooter {

    public static void main(String[] args) {
        String inputPath = "path/to/input.docx";
        String outputPath = "path/to/output.docx";

        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputPath));

            // Remove hyperlinks in all footer sections
            removeHyperlinksFromFooters(wordMLPackage);

            // Save the updated document
            wordMLPackage.save(new File(outputPath));

            System.out.println("Hyperlinks removed successfully, text restored, and saved to: " + outputPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void removeHyperlinksFromFooters(WordprocessingMLPackage wordMLPackage) throws Docx4JException {
        // Access footer parts through relationships
        List<Relationship> footerRelationships = wordMLPackage.getMainDocumentPart()
                                                              .getRelationshipsPart()
                                                              .getRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer");

        for (Relationship rel : footerRelationships) {
            FooterPart footerPart = (FooterPart) wordMLPackage.getParts().get(new PartName(rel.getTarget()));

            if (footerPart != null) {
                List<Object> footerContent = footerPart.getContent();

                for (Object obj : footerContent) {
                    if (obj instanceof JAXBElement) {
                        JAXBElement<?> element = (JAXBElement<?>) obj;
                        Object value = element.getValue();

                        if (value instanceof P) {
                            removeHyperlinksFromParagraph((P) value);
                        }
                    }
                }
            }
        }
    }

    private static void removeHyperlinksFromParagraph(P paragraph) {
        List<Object> children = paragraph.getContent();

        for (int i = 0; i < children.size(); i++) {
            Object child = children.get(i);
            if (child instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) child;
                Object value = element.getValue();

                if (value instanceof P.Hyperlink) {
                    P.Hyperlink hyperlink = (P.Hyperlink) value;

                    // Extract the text from the hyperlink content
                    String hyperlinkText = extractText(hyperlink.getContent());

                    // Replace hyperlink with text
                    R plainTextRun = createPlainTextRun(hyperlinkText, hyperlink);
                    paragraph.getContent().set(i, plainTextRun);
                }
            }
        }
    }

    private static String extractText(List<Object> content) {
        StringBuilder textBuilder = new StringBuilder();
        for (Object child : content) {
            if (child instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) child;
                if (element.getValue() instanceof R) {
                    R run = (R) element.getValue();
                    for (Object runContent : run.getContent()) {
                        if (runContent instanceof Text) {
                            textBuilder.append(((Text) runContent).getValue());
                        }
                    }
                }
            }
        }
        return textBuilder.toString();
    }

    private static R createPlainTextRun(String text, P.Hyperlink hyperlink) {
        ObjectFactory factory = new ObjectFactory();
        R newRun = factory.createR();

        // Preserve formatting if available
        if (!hyperlink.getContent().isEmpty() && hyperlink.getContent().get(0) instanceof JAXBElement) {
            JAXBElement<?> firstElement = (JAXBElement<?>) hyperlink.getContent().get(0);
            if (firstElement.getValue() instanceof R) {
                R existingRun = (R) firstElement.getValue();
                newRun.setRPr(existingRun.getRPr());
            }
        }

        // Add text content back without the hyperlink
        Text textElement = factory.createText();
        textElement.setValue(text);
        newRun.getContent().add(textElement);

        return newRun;
    }
}

----------
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.List;

public class RemoveHyperlinksInFooter {

    public static void main(String[] args) {
        String inputPath = "path/to/input.docx";
        String outputPath = "path/to/output.docx";

        try {
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputPath));

            // Remove hyperlinks in footers
            removeHyperlinksFromFooters(wordMLPackage);

            // Save the updated file
            wordMLPackage.save(new File(outputPath));

            System.out.println("Hyperlinks in footers removed successfully and saved to: " + outputPath);
        } catch (Docx4JException e) {
            e.printStackTrace();
        }
    }

    private static void removeHyperlinksFromFooters(WordprocessingMLPackage wordMLPackage) throws Docx4JException {
        // Get all footer parts
        List<FooterPart> footerParts = wordMLPackage.getDocumentModel().getParts().getPartsOfType(FooterPart.class);

        for (FooterPart footerPart : footerParts) {
            List<Object> footerContent = footerPart.getJaxbElement().getContent();

            for (Object obj : footerContent) {
                if (obj instanceof JAXBElement) {
                    JAXBElement<?> element = (JAXBElement<?>) obj;
                    Object value = element.getValue();

                    if (value instanceof P) {
                        removeHyperlinksFromParagraph((P) value);
                    }
                }
            }
        }
    }

    private static void removeHyperlinksFromParagraph(P paragraph) {
        List<Object> children = paragraph.getContent();

        for (int i = 0; i < children.size(); i++) {
            Object child = children.get(i);
            if (child instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) child;
                Object value = element.getValue();

                if (value instanceof P.Hyperlink) {
                    // Extract the hyperlink content as plain text
                    R combinedRun = extractTextRuns(((P.Hyperlink) value).getContent());

                    // Replace the hyperlink with plain text
                    paragraph.getContent().set(i, combinedRun);
                }
            }
        }
    }

    private static R extractTextRuns(List<Object> content) {
        ObjectFactory factory = new ObjectFactory();
        R combinedRun = factory.createR();

        for (Object child : content) {
            if (child instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) child;
                if (element.getValue() instanceof R) {
                    R run = (R) element.getValue();
                    combinedRun.getContent().addAll(run.getContent());
                    combinedRun.setRPr(run.getRPr());
                }
            }
        }
        return combinedRun;
    }
}

------

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.ArrayList;
import java.util.List;

public class RemoveHyperlinksDocx4jJava11 {

    public static void main(String[] args) {
        String inputPath = "path/to/input.docx";
        String outputPath = "path/to/output.docx";

        try {
            // Load the Word document
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputPath));

            // Process all paragraphs and remove hyperlinks
            removeHyperlinks(wordMLPackage);

            // Save the updated document
            wordMLPackage.save(new File(outputPath));

            System.out.println("Hyperlinks removed successfully and saved to: " + outputPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void removeHyperlinks(WordprocessingMLPackage wordMLPackage) {
        List<Object> paragraphs = wordMLPackage.getMainDocumentPart().getContent();

        for (Object obj : paragraphs) {
            if (obj instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) obj;
                Object value = element.getValue();

                if (value instanceof P) {
                    P paragraph = (P) value;
                    removeHyperlinksFromParagraph(paragraph);
                }
            }
        }
    }

    private static void removeHyperlinksFromParagraph(P paragraph) {
        List<Object> children = new ArrayList<>(paragraph.getContent());
        List<Object> updatedContent = new ArrayList<>();

        for (Object child : children) {
            updatedContent.add(processHyperlink(child));
        }

        paragraph.getContent().clear();
        paragraph.getContent().addAll(updatedContent);
    }

    private static Object processHyperlink(Object obj) {
        if (obj instanceof JAXBElement) {
            JAXBElement<?> element = (JAXBElement<?>) obj;
            Object value = element.getValue();

            if (value instanceof P.Hyperlink) {
                return extractTextRuns(((P.Hyperlink) value).getContent());
            }
        }
        return obj;
    }

    private static R extractTextRuns(List<Object> content) {
        ObjectFactory factory = new ObjectFactory();
        R combinedRun = factory.createR();

        for (Object child : content) {
            if (child instanceof JAXBElement) {
                JAXBElement<?> element = (JAXBElement<?>) child;
                if (element.getValue() instanceof R) {
                    R run = (R) element.getValue();
                    combinedRun.getContent().addAll(run.getContent());
                    combinedRun.setRPr(run.getRPr());
                }
            }
        }
        return combinedRun;
    }
}

-----------
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;
import java.io.File;
import java.util.List;
import java.util.stream.Collectors;

public class RemoveHyperlinksDocx4j {

    public static void main(String[] args) {
        String inputPath = "path/to/input.docx";
        String outputPath = "path/to/output.docx";

        try {
            // Load the Word document
            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new File(inputPath));

            // Process all paragraphs and remove hyperlinks
            removeHyperlinks(wordMLPackage);

            // Save the updated document
            wordMLPackage.save(new File(outputPath));

            System.out.println("Hyperlinks removed successfully and saved to: " + outputPath);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void removeHyperlinks(WordprocessingMLPackage wordMLPackage) {
        List<Object> paragraphs = wordMLPackage.getMainDocumentPart().getContent();

        for (Object obj : paragraphs) {
            if (obj instanceof JAXBElement) {
                Object value = ((JAXBElement<?>) obj).getValue();

                if (value instanceof P paragraph) {
                    removeHyperlinksFromParagraph(paragraph);
                }
            }
        }
    }

    private static void removeHyperlinksFromParagraph(P paragraph) {
        List<Object> children = paragraph.getContent();

        // Collect non-hyperlinked elements and keep text nodes
        List<Object> updatedContent = children.stream()
                .map(RemoveHyperlinksDocx4j::processHyperlink)
                .collect(Collectors.toList());

        paragraph.getContent().clear();
        paragraph.getContent().addAll(updatedContent);
    }

    private static Object processHyperlink(Object obj) {
        if (obj instanceof JAXBElement) {
            JAXBElement<?> element = (JAXBElement<?>) obj;
            Object value = element.getValue();

            // Check for hyperlink tags
            if (value instanceof org.docx4j.wml.P.Hyperlink hyperlink) {
                return extractTextRuns(hyperlink.getContent());
            }
        }
        return obj;
    }

    @SuppressWarnings("unchecked")
    private static Object extractTextRuns(List<Object> content) {
        R combinedRun = new ObjectFactory().createR();

        for (Object child : content) {
            if (child instanceof JAXBElement) {
                JAXBElement<?> textElement = (JAXBElement<?>) child;
                if (textElement.getValue() instanceof R run) {
                    combinedRun.getContent().addAll(run.getContent());
                    combinedRun.setRPr(run.getRPr());
                }
            }
        }
        return combinedRun;
    }
}

-------------
<w:hyperlink r:id="rId1" w:history="1">
						<w:r w:rsidR="004522BB" w:rsidRPr="002606AA">
							<w:rPr>
								<w:rStyle w:val="Hyperlink"/>
								<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
								<w:b/>
								<w:sz w:val="16"/>
								<w:szCs w:val="16"/>
							</w:rPr>
							<w:t>www.bluecrossma.com</w:t>
						</w:r>
					</w:hyperlink>
-----------------


import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class RemoveHyperlinksJava11 {

    public static void main(String[] args) {
        String inputDocxPath = "path/to/input.docx";
        String outputDocxPath = "path/to/output.docx";

        try (FileInputStream fis = new FileInputStream(inputDocxPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            removeHyperlinks(document);

            try (FileOutputStream fos = new FileOutputStream(outputDocxPath)) {
                document.write(fos);
            }

            System.out.println("Hyperlinks removed and saved to: " + outputDocxPath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void removeHyperlinks(XWPFDocument document) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs == null) continue;

            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                // Check if it's a hyperlink run using casting instead of instanceof patterns
                if (run != null && run.getClass().equals(XWPFHyperlinkRun.class)) {
                    XWPFHyperlinkRun hyperlinkRun = (XWPFHyperlinkRun) run;

                    // Extract hyperlink text
                    String text = hyperlinkRun.text();

                    // Replace with a regular run
                    XWPFRun newRun = paragraph.insertNewRun(i);
                    newRun.setText(text);

                    // Copy styles to the new run
                    copyRunStyle(hyperlinkRun, newRun);

                    // Remove old hyperlink run
                    paragraph.removeRun(i + 1);
                }
            }
        }
    }

    private static void copyRunStyle(XWPFRun source, XWPFRun target) {
        target.setBold(source.isBold());
        target.setItalic(source.isItalic());
        target.setUnderline(source.getUnderline());
        target.setFontSize(source.getFontSize());
        target.setFontFamily(source.getFontFamily());
        target.setColor(source.getColor());
    }
}

------------

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class RemoveHyperlinks {

    public static void main(String[] args) {
        String inputDocxPath = "path/to/input.docx";
        String outputDocxPath = "path/to/output.docx";

        try (FileInputStream fis = new FileInputStream(inputDocxPath);
             XWPFDocument document = new XWPFDocument(fis)) {

            removeHyperlinks(document);

            try (FileOutputStream fos = new FileOutputStream(outputDocxPath)) {
                document.write(fos);
            }

            System.out.println("Hyperlinks removed and saved to: " + outputDocxPath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void removeHyperlinks(XWPFDocument document) {
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs == null) continue;

            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                if (run instanceof XWPFHyperlinkRun hyperlinkRun) {
                    String text = hyperlinkRun.text();  // Extract hyperlink text
                    XWPFRun newRun = paragraph.insertNewRun(i);  // Replace with a regular run
                    newRun.setText(text);

                    copyRunStyle(hyperlinkRun, newRun);
                    paragraph.removeRun(i + 1);  // Remove old hyperlink run
                }
            }
        }
    }

    private static void copyRunStyle(XWPFRun source, XWPFRun target) {
        target.setBold(source.isBold());
        target.setItalic(source.isItalic());
        target.setUnderline(source.getUnderline());
        target.setFontSize(source.getFontSize());
        target.setFontFamily(source.getFontFamily());
        target.setColor(source.getColor());
    }
}


--------


import com.itextpdf.io.image.ImageData;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;

public class DocxToPdfWithImages {
    public static void main(String[] args) {
        String docxPath = "path/to/input.docx";
        String pdfPath = "path/to/output.pdf";

        try (XWPFDocument docx = new XWPFDocument(new FileInputStream(docxPath));
             PdfWriter writer = new PdfWriter(pdfPath);
             PdfDocument pdf = new PdfDocument(writer);
             Document document = new Document(pdf)) {

            for (XWPFParagraph paragraph : docx.getParagraphs()) {
                // Add text from DOCX
                String text = paragraph.getText();
                if (!text.isEmpty()) {
                    document.add(new Paragraph(text));
                }
            }

            // Handle images in the DOCX
            List<XWPFPictureData> pictures = docx.getAllPictures();
            for (XWPFPictureData pictureData : pictures) {
                byte[] imageBytes = pictureData.getData();
                String extension = pictureData.suggestFileExtension();
                
                if (extension.equalsIgnoreCase("jpeg") || extension.equalsIgnoreCase("png")) {
                    ImageData imageData = ImageDataFactory.create(imageBytes);
                    Image image = new Image(imageData);
                    image.setAutoScale(true); // Automatically resize to fit
                    document.add(image);
                }
            }

            System.out.println("Conversion to PDF completed.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}




<dependency>
    <groupId>fr.opensagres.xdocreport</groupId>
    <artifactId>xdocreport-document-injector</artifactId>
    <version>2.0.2</version>
</dependency>

<dependency>
    <groupId>fr.opensagres.xdocreport</groupId>
    <artifactId>fr.opensagres.xdocreport.converter.docx.xwpf</artifactId>
    <version>2.0.2</version>
</dependency>

<dependency>
    <groupId>fr.opensagres.xdocreport</groupId>
    <artifactId>fr.opensagres.xdocreport.converter.pdf.itext</artifactId>
    <version>2.0.2</version>
</dependency>



import fr.opensagres.xdocreport.converter.ConverterTypeTo;
import fr.opensagres.xdocreport.converter.Options;
import fr.opensagres.xdocreport.converter.XDocConverterException;
import fr.opensagres.xdocreport.converter.pdf.PDFViaITextOptions;
import fr.opensagres.xdocreport.core.document.DocumentKind;
import fr.opensagres.xdocreport.document.IConverter;
import fr.opensagres.xdocreport.document.IWContext;
import fr.opensagres.xdocreport.document.registry.DocumentKindRegistry;
import fr.opensagres.xdocreport.template.TemplateEngineKind;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;

public class DocxToPdfConverterXDocReport {

    public static void main(String[] args) {
        String docxPath = "path/to/your/input.docx";
        String pdfPath = "path/to/your/output.pdf";

        try (InputStream docxInputStream = new FileInputStream(docxPath);
             OutputStream pdfOutputStream = new FileOutputStream(pdfPath)) {

            // Create XWPFDocument instance
            XWPFDocument document = new XWPFDocument(docxInputStream);

            // Prepare XDocReport Converter options
            Options options = Options.getFrom(DocumentKind.DOCX)
                                      .to(ConverterTypeTo.PDF)
                                      .via(PDFViaITextOptions.create().compress());

            IConverter converter = DocumentKindRegistry.getRegistry().getConverter(options);

            // Perform the conversion
            converter.convert(document, pdfOutputStream, IWContext.create());
            System.out.println("Conversion to PDF completed.");

        } catch (IOException | XDocConverterException e) {
            e.printStackTrace();
        }
    }
}

------------------

    <dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi</artifactId>
			<version>4.0.1</version>
		</dependency>
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-ooxml</artifactId>
			<version>4.0.1</version>
		</dependency>
		<dependency>
			<groupId>fr.opensagres.xdocreport</groupId>
			<artifactId>fr.opensagres.poi.xwpf.converter.pdf</artifactId>
			<version>2.0.2</version>
		</dependency>

	----------

	// Handle image extraction in complex documents
            options.setExtractor(new PdfImageExtractor() {
                @Override
                public void extractImage(byte[] imageData, int imageType, float width, float height) {
                    System.out.println("Extracting image with width: " + width + " and height: " + height);
                }
            });

------------

import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.List;

public class DocxToPdfConverter {

    public static void main(String[] args) {
        String docxPath = "path/to/your/input.docx";
        String pdfPath = "path/to/your/output.pdf";

        try (FileInputStream docxInputStream = new FileInputStream(docxPath);
             FileOutputStream pdfOutputStream = new FileOutputStream(pdfPath)) {

            // Load the DOCX file with POI
            XWPFDocument document = new XWPFDocument(docxInputStream);

            // Ensure proper handling of complex images
            extractImages(document);

            // Set up PDF conversion options
            PdfOptions options = PdfOptions.create();

            // Perform the conversion
            PdfConverter.getInstance().convert(document, pdfOutputStream, options);

            System.out.println("DOCX file successfully converted to PDF.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void extractImages(XWPFDocument document) {
        List<XWPFPictureData> pictures = document.getAllPictures();
        if (pictures.isEmpty()) {
            System.out.println("No embedded images found in the DOCX.");
        } else {
            System.out.println("Found " + pictures.size() + " images:");
            for (XWPFPictureData pictureData : pictures) {
                String imageType = pictureData.suggestFileExtension();
                System.out.println("Image Type: " + imageType);
                try (FileOutputStream imageOut = new FileOutputStream("output_image_" + pictureData.getPackagePart().getPartName().getName())) {
                    imageOut.write(pictureData.getData());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

}
------------


import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DocxToPdfWithImages {

    public static void main(String[] args) {
        String inputDocxPath = "path/to/input.docx";
        String outputPdfPath = "path/to/output.pdf";

        try (FileInputStream inputStream = new FileInputStream(inputDocxPath);
             FileOutputStream pdfOutputStream = new FileOutputStream(outputPdfPath)) {

            // Load the document
            XWPFDocument document = new XWPFDocument(inputStream);

            // Extract images and store with their paragraph indices
            Map<Integer, byte[]> imageMap = extractImages(document);

            // Reinsert the images into the correct places
            reinsertImages(document, imageMap);

            // Convert the modified document to PDF
            PdfOptions options = PdfOptions.create();
            PdfConverter.getInstance().convert(document, pdfOutputStream, options);

            System.out.println("DOCX successfully converted to PDF with images.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static Map<Integer, byte[]> extractImages(XWPFDocument document) {
        Map<Integer, byte[]> imageMap = new HashMap<>();
        int pictureIndex = 0;

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null) {
                for (XWPFRun run : runs) {
                    List<XWPFPicture> pictures = run.getEmbeddedPictures();
                    for (XWPFPicture picture : pictures) {
                        XWPFPictureData pictureData = picture.getPictureData();
                        imageMap.put(pictureIndex++, pictureData.getData());
                        System.out.println("Extracted image at index: " + pictureIndex);
                    }
                }
            }
        }
        return imageMap;
    }

    private static void reinsertImages(XWPFDocument document, Map<Integer, byte[]> imageMap) {
        int imageIndex = 0;

        for (XWPFParagraph paragraph : document.getParagraphs()) {
            List<XWPFRun> runs = paragraph.getRuns();
            if (runs != null && !imageMap.isEmpty() && imageIndex < imageMap.size()) {
                for (XWPFRun run : runs) {
                    try {
                        byte[] imageData = imageMap.get(imageIndex++);
                        if (imageData != null) {
                            run.addPicture(new ByteArrayInputStream(imageData),
                                    XWPFDocument.PICTURE_TYPE_PNG, "image" + imageIndex + ".png",
                                    200, 150); // Width and height adjustments
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }
}
----------
	
    
