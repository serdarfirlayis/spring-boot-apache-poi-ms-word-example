package com.serdarfirlayis.poimswordexample.service;

import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

@Service
public class PoiMsWordExampleService {

    public void createWordDocument() {

        // creating a document
        XWPFDocument document = new XWPFDocument();

        // creating a paragraph
        XWPFParagraph paragraph = document.createParagraph();

        // creating a run
        XWPFRun run = paragraph.createRun();

        // setting text to run
        run.setText("Hello World! Poi MS Word Example");
        run.setBold(true);
        run.setFontSize(20);

        // creating a header
        XWPFHeader header = document.createHeader(HeaderFooterType.DEFAULT);
        XWPFParagraph paragraphHeader = header.createParagraph();
        XWPFRun runHeader = paragraphHeader.createRun();
        runHeader.setText("Header");
        runHeader.setBold(true);
        runHeader.setFontSize(24);

        // creating a footer
        XWPFFooter footer = document.createFooter(HeaderFooterType.DEFAULT);
        XWPFParagraph paragraphFooter = footer.createParagraph();
        XWPFRun runFooter = paragraphFooter.createRun();
        runFooter.setText("Footer");
        runFooter.setBold(true);
        runFooter.setFontSize(24);

        // writing to a file
        try (FileOutputStream out = new FileOutputStream(new File("src/main/resources/wordDocument.docx"))) {
            document.write(out);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
