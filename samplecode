
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
