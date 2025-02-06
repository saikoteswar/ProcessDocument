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
	
    
