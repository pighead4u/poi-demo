package com.ashideng.report;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

/**
 * @Description:
 * @Author: zhuhuanhuan@shunjiantech.cn
 * @Date: 2022/11/28 上午9:34
 * @Version: 1.0.0
 **/
public class TestReadReport {

    public void readReports(String path) {
        try(XWPFDocument doc = new XWPFDocument(new FileInputStream(path))) {
            List<XWPFParagraph> paragraphs = doc.getParagraphs();
            XWPFParagraph first = paragraphs.get(0);
            String text = first.getParagraphText();
            System.out.println(text);

            List<XWPFTable> tables = doc.getTables();
            XWPFTable first_table = tables.get(0);
            int width = first_table.getWidth();
            System.out.println("width: " + width);
            text = first_table.getText();
            System.out.println(text);
            int row_num = first_table.getNumberOfRows();
            System.out.println(row_num);

            List<XWPFTableRow> first_table_rows = first_table.getRows();
            first_table_rows.forEach(item -> {
                int height = item.getHeight();
                System.out.println("row height:" + height);
            });

            XWPFTable second_table = tables.get(1);
            width = second_table.getWidth();
            System.out.println("222width: " + width);
            text = second_table.getText();
            System.out.println(text);
            row_num = second_table.getNumberOfRows();
            System.out.println(row_num);

            List<XWPFTableRow> second_table_rows = second_table.getRows();
            second_table_rows.forEach(item -> {
                int height = item.getHeight();
                System.out.println("222row height:" + height);
            });


        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
