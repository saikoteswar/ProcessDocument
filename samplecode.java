
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
    
