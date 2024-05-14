package com.excel.PocExcel;

import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.kernel.colors.DeviceGray;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.Rectangle;
import com.itextpdf.kernel.pdf.*;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.layout.Canvas;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.VerticalAlignment;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeUtils;
import org.jodconverter.local.office.LocalOfficeManager;
import org.jodconverter.local.office.OfficeConnectionException;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.IOException;
import java.rmi.ConnectException;

import org.jodconverter.core.*;
import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.document.DocumentFamily;
import org.jodconverter.core.document.DocumentFormat;
import org.jodconverter.local.*;
import org.jodconverter.local.office.OfficeConnection;
//import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
//import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;

@SpringBootApplication
public class ExcelConversion {

   public OfficeConnection connection;

    public static void main(String[] args) throws DocumentException, IOException, OfficeException {

        SpringApplication.run(ExcelConversion.class, args);
        System.out.println(" iniciando ...");
        convertirExcelPDF();
        adicionarSelloAgua("Sello de agua");
        encryptarPDF();
    }

    public static void convertirExcelPDF() throws OfficeException {

        //File inputFile = new File("\\DATA\\FV3800.xls");
        File inputFile = new File("\\DATA\\FV3800IMAGEN.xls");
        File outputFile = new File("\\DATA\\FV3800LF.pdf");

        final LocalOfficeManager officeManager = LocalOfficeManager.install();
        try {

            // Start an office process and connect to the started instance (on port 2002).
            officeManager.start();

            // Convert
            JodConverter
                    .convert(inputFile)
                    .to(outputFile)
                    .execute();
        } finally {
            // Stop the office process
            OfficeUtils.stopQuietly(officeManager);
        }
    }

    static void adicionarSelloAgua(String watermarkText){

        try {
            PdfReader reader = new PdfReader(new File("\\DATA\\FV3800LF.pdf"));
            com.itextpdf.kernel.pdf.PdfWriter writer = new com.itextpdf.kernel.pdf.PdfWriter("\\DATA\\FV3800SA.pdf");

            PdfDocument pdfDoc = new PdfDocument(reader,writer);
            int numPages = pdfDoc.getNumberOfPages();


            // Cargar la fuente Helvetica
            com.itextpdf.kernel.font.PdfFont helvetica = PdfFontFactory.createFont(StandardFonts.HELVETICA);

            // Agregar sello de agua a cada página del documento
            for (int i = 1; i <= numPages; i++) {
                PdfPage page = pdfDoc.getPage(i);
                PdfCanvas pdfCanvas = new PdfCanvas(page);
                Rectangle pageSize = page.getPageSize();
                Canvas canvas = new Canvas(pdfCanvas, pageSize);
                canvas.setFont(helvetica);
                canvas.setFontSize(50);
                canvas.setFontColor(new DeviceGray(0.9f)); // Color gris claro
                canvas.showTextAligned(new Paragraph(watermarkText), pageSize.getWidth() / 2, pageSize.getHeight() / 2, i, TextAlignment.CENTER, VerticalAlignment.MIDDLE, 45);
            }

            // Cerrar el documento
            pdfDoc.close();
            System.out.println("Sello de agua agregado correctamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    static void encryptarPDF()  throws IOException, DocumentException{

        try {
            Document document = new Document();
            //PdfDocument documentPDF = new PdfDocument(new PdfReader("D:\\DATA\\FV3800SA.pdf"), new PdfWriter("D:\\DATA\\FV3800SAE.pdf",new WriterProperties().setStandardEncryption(Encoding.UTF8.GetBytes("Hello"), Encoding.UTF8.GetBytes("Hello"), com.itextpdf.text.pdf.PdfWriter.ALLOW_PRINTING, com.itextpdf.text.pdf.PdfWriter.STANDARD_ENCRYPTION_128)));
            //PdfDocument documentPDF = new PdfDocument(new PdfReader("D:\\DATA\\FV3800SA.pdf"), new PdfWriter("D:\\DATA\\FV3800SAE.pdf",new WriterProperties().setStandardEncryption("my-owner-password".getBytes(), "my-user-password".getBytes(), com.itextpdf.text.pdf.PdfWriter.ALLOW_PRINTING | com.itextpdf.text.pdf.PdfWriter.ALLOW_COPY , com.itextpdf.text.pdf.PdfWriter.ENCRYPTION_AES_256)));
            PdfDocument documentPDF = new PdfDocument(new PdfReader("\\DATA\\FV3800SA.pdf"), new PdfWriter("\\DATA\\FV3800SAE.pdf",new WriterProperties().setStandardEncryption("my-owner-password".getBytes(), "my-user-password".getBytes(), 0 , com.itextpdf.text.pdf.PdfWriter.ENCRYPTION_AES_256)));
            documentPDF.close();

            System.out.println("Encriptación correctamente.");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

}
