package com.ashideng.report.template;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @Description:
 * @Author: zhuhuanhuan@shunjiantech.cn
 * @Date: 2022/11/28 下午1:55
 * @Version: 1.0.0
 **/
public class WordInsertContentInTable {
    static void setText(XWPFTableCell cell, String text) {
        String[] lines = text.split("\n");
        List<XWPFParagraph> paragraphs = cell.getParagraphs();
        for (int i = 0; i < lines.length; i++) {
            String line = lines[i];
            XWPFParagraph paragraph = null;
            if (paragraphs.size() > i) paragraph = paragraphs.get(i);
            if (paragraph == null) paragraph = cell.addParagraph();
            XWPFRun run = null;
            if (paragraph.getRuns().size() > 0) run = paragraph.getRuns().get(0);
            if (run == null) run = paragraph.createRun();
            run.setText(line, 0);
        }
        for (int i = paragraphs.size()-1; i >= lines.length; i--) {
            cell.removeParagraph(i);
        }
    }

    static void insertContentInTable(XWPFTable table, List<POJO> listOfPOJOs) throws Exception {
        XWPFTableRow titleRowTemplate = table.getRow(0);
        if (titleRowTemplate == null) throw new Exception("Table template does not match: No title row.");
        if (titleRowTemplate.getTableCells().size() != 1) throw new Exception("Table template does not match: Wrong title row column count.");
        XWPFTableRow subTitleRowTemplate = table.getRow(1);
        if (subTitleRowTemplate == null) throw new Exception("Table template does not match: No sub title row.");
        if (subTitleRowTemplate.getTableCells().size() != 2) throw new Exception("Table template does not match: Wrong sub title row column count.");
        XWPFTableRow contentRowTemplate = table.getRow(2);
        if (contentRowTemplate == null) throw new Exception("Table template does not match: No content row.");
        if (contentRowTemplate.getTableCells().size() != 2) throw new Exception("Table template does not match: Wrong content row column count.");

        XWPFTableRow titleRow = titleRowTemplate;
        XWPFTableRow subTitleRow = subTitleRowTemplate;
        XWPFTableRow contentRow = contentRowTemplate;
        XWPFTableCell cell;
        for (int i = 0; i < listOfPOJOs.size(); i++) {
            POJO pojo = listOfPOJOs.get(i);
            if (i > 0) {
                titleRow = new XWPFTableRow((CTRow)titleRowTemplate.getCtRow().copy(), table);
                subTitleRow = new XWPFTableRow((CTRow)subTitleRowTemplate.getCtRow().copy(), table);
                contentRow = new XWPFTableRow((CTRow)contentRowTemplate.getCtRow().copy(), table);
            }
            String titleRowText = pojo.getTitleRowText();
            cell = titleRow.getCell(0);
            setText(cell, titleRowText);
            String subTitleRowLeftText = pojo.getSubTitleRowLeftText();
            String subTitleRowLeftColor = pojo.getSubTitleRowLeftColor();
            String subTitleRowRightText = pojo.getSubTitleRowRightText();
            cell = subTitleRow.getCell(0);
            setText(cell,subTitleRowLeftText);
            cell.setColor(subTitleRowLeftColor);
            cell = subTitleRow.getCell(1);
            setText(cell,subTitleRowRightText);
            String contentRowLeftText = pojo.getContentRowLeftText();
            String contentRowRightText = pojo.getContentRowRightText();
            cell = contentRow.getCell(0);
            setText(cell, contentRowLeftText);
            cell = contentRow.getCell(1);
            setText(cell, contentRowRightText);
            if (i > 0) {
                table.addRow(titleRow);
                table.addRow(subTitleRow);
                table.addRow(contentRow);
            }
        }
    }

    public static void main(String[] args) throws Exception {

        List<POJO> listOfPOJOs = new ArrayList<POJO>();
        listOfPOJOs.add(new POJO("Title row text 1",
                "Sub title row left text 1", "FF0000", "Sub title row right text 1\nSub title row right text 1\nSub title row right text 1",
                "Content row left text 1\nContent row left text 1\nContent row left text 1",
                "Content row right text 1\nContent row right text 1\nContent row right text 1"));
        listOfPOJOs.add(new POJO("Title row text 2",
                "Sub title row left text 2", "00FF00", "Sub title row right text 2\nSub title row right text 2",
                "Content row left text 2\nContent row left text 2",
                "Content row right text 2\nContent row right text 2"));
        listOfPOJOs.add(new POJO("Title row text 3",
                "Sub title row left text 3", "0000FF", "Sub title row right text 3",
                "Content row left text 3",
                "Content row right text 3"));

        XWPFDocument document = new XWPFDocument(new FileInputStream("./templates/wordtemplate.docx"));

        XWPFTable table = document.getTableArray(0);

        insertContentInTable(table, listOfPOJOs);

        FileOutputStream out = new FileOutputStream("./WordResult.docx");
        document.write(out);
        out.close();
        document.close();
    }

    static class POJO {
        private String titleRowText;
        private String subTitleRowLeftText;
        private String subTitleRowLeftColor;
        private String subTitleRowRightText;
        private String contentRowLeftText;
        private String contentRowRightText;
        public POJO ( String titleRowText,
                      String subTitleRowLeftText,
                      String subTitleRowLeftColor,
                      String subTitleRowRightText,
                      String contentRowLeftText,
                      String contentRowRightText ) {
            this.titleRowText = titleRowText;
            this.subTitleRowLeftText = subTitleRowLeftText;
            this.subTitleRowLeftColor = subTitleRowLeftColor;
            this.subTitleRowRightText = subTitleRowRightText;
            this.contentRowLeftText = contentRowLeftText;
            this.contentRowRightText = contentRowRightText;
        }
        public String getTitleRowText() {
            return this.titleRowText;
        }
        public String getSubTitleRowLeftText() {
            return this.subTitleRowLeftText;
        }
        public String getSubTitleRowLeftColor() {
            return this.subTitleRowLeftColor;
        }
        public String getSubTitleRowRightText() {
            return this.subTitleRowRightText;
        }
        public String getContentRowLeftText() {
            return this.contentRowLeftText;
        }
        public String getContentRowRightText() {
            return this.contentRowRightText;
        }
    }
}
