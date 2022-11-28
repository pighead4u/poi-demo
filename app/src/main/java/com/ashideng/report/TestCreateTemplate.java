package com.ashideng.report;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @Description:
 * @Author: zhuhuanhuan@shunjiantech.cn
 * @Date: 2022/11/28 上午8:38
 * @Version: 1.0.0
 **/
public class TestCreateTemplate {
//    private XWPFDocument docx;
//
//    public TestCreateTemplate(String path) {
//        try {
//            FileInputStream file = new FileInputStream(path);
//            docx = new XWPFDocument(file);
//        } catch (FileNotFoundException e) {
//            throw new RuntimeException(e);
//        } catch (IOException e) {
//            throw new RuntimeException(e);
//        }
//    }

    public void testCreate() {
        try (XWPFDocument doc = new XWPFDocument()) {

            XWPFParagraph p1 = doc.createParagraph();
            p1.setAlignment(ParagraphAlignment.CENTER);
            p1.setBorderBottom(Borders.DOUBLE);
            p1.setBorderTop(Borders.DOUBLE);

            p1.setBorderRight(Borders.DOUBLE);
            p1.setBorderLeft(Borders.DOUBLE);
            p1.setBorderBetween(Borders.SINGLE);

            p1.setVerticalAlignment(TextAlignment.TOP);

            XWPFRun r1 = p1.createRun();
            r1.setBold(true);
            r1.setText("The quick brown fox");
            r1.setBold(true);
            r1.setFontFamily("Courier");
            r1.setUnderline(UnderlinePatterns.DOT_DOT_DASH);
            r1.setTextPosition(100);

//            XWPFParagraph p2 = doc.createParagraph();
//            p2.setAlignment(ParagraphAlignment.RIGHT);
//
//            //BORDERS
//            p2.setBorderBottom(Borders.DOUBLE);
//            p2.setBorderTop(Borders.DOUBLE);
//            p2.setBorderRight(Borders.DOUBLE);
//            p2.setBorderLeft(Borders.DOUBLE);
//            p2.setBorderBetween(Borders.SINGLE);
//
//            XWPFRun r2 = p2.createRun();
//            r2.setText("jumped over the lazy dog");
//            r2.setStrikeThrough(true);
//            r2.setFontSize(20);
//
//            XWPFRun r3 = p2.createRun();
//            r3.setText("and went away");
//            r3.setStrikeThrough(true);
//            r3.setFontSize(20);
//            r3.setSubscript(VerticalAlign.SUPERSCRIPT);
//
//            // hyperlink
//            XWPFHyperlinkRun hyperlink = p2.insertNewHyperlinkRun(0, "http://poi.apache.org/");
//            hyperlink.setUnderline(UnderlinePatterns.SINGLE);
//            hyperlink.setColor("0000ff");
//            hyperlink.setText("Apache POI");
//
//            XWPFParagraph p3 = doc.createParagraph();
//            p3.setWordWrapped(true);
//            p3.setPageBreak(true);
//
//            //p3.setAlignment(ParagraphAlignment.DISTRIBUTE);
//            p3.setAlignment(ParagraphAlignment.BOTH);
//            p3.setSpacingBetween(15, LineSpacingRule.EXACT);
//
//            p3.setIndentationFirstLine(600);
//
//
//            XWPFRun r4 = p3.createRun();
//            r4.setTextPosition(20);
//            r4.setText("To be, or not to be: that is the question: "
//                    + "Whether 'tis nobler in the mind to suffer "
//                    + "The slings and arrows of outrageous fortune, "
//                    + "Or to take arms against a sea of troubles, "
//                    + "And by opposing end them? To die: to sleep; ");
//            r4.addBreak(BreakType.PAGE);
//            r4.setText("No more; and by a sleep to say we end "
//                    + "The heart-ache and the thousand natural shocks "
//                    + "That flesh is heir to, 'tis a consummation "
//                    + "Devoutly to be wish'd. To die, to sleep; "
//                    + "To sleep: perchance to dream: ay, there's the rub; "
//                    + ".......");
//            r4.setItalic(true);
////This would imply that this break shall be treated as a simple line break, and break the line after that word:
//
//            XWPFRun r5 = p3.createRun();
//            r5.setTextPosition(-10);
//            r5.setText("For in that sleep of death what dreams may come");
//            r5.addCarriageReturn();
//            r5.setText("When we have shuffled off this mortal coil, "
//                    + "Must give us pause: there's the respect "
//                    + "That makes calamity of so long life;");
//            r5.addBreak();
//            r5.setText("For who would bear the whips and scorns of time, "
//                    + "The oppressor's wrong, the proud man's contumely,");
//
//            r5.addBreak(BreakClear.ALL);
//            r5.setText("The pangs of despised love, the law's delay, "
//                    + "The insolence of office and the spurns " + ".......");

            try (FileOutputStream out = new FileOutputStream("simple.docx")) {
                doc.write(out);
            } catch (FileNotFoundException e) {
                throw new RuntimeException(e);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
