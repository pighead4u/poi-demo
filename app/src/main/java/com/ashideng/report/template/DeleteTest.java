package com.ashideng.report.template;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author zhuhuanhuan@shunjiantech.cn
 * @version 1.0.0
 * @description
 * @create 2022/12/9 下午2:34
 **/
public class DeleteTest {

    // 定位试验table的正则
    private final String PATTERN = "(?<=\\{\\{\\^)[^\\}\\}]+";

    public void buildNewTemplate(String path, Set<String> excludeTable) {
        try {
            XWPFDocument document = new XWPFDocument(new FileInputStream(path));
            List<XWPFTable> tables = document.getTables();

            Set<XWPFTable> deleteTables = new HashSet<>();
            Pattern pattern = Pattern.compile(PATTERN);
            Set<Integer> keepParagraphs = new HashSet<>();

            for (int i = 6; i < tables.size(); i++) {
                XWPFTable table = tables.get(i);
                String content = table.getText();
                Matcher matcher = pattern.matcher(content);
                if (matcher.find()) {
                    String extractContent = matcher.group();
                    if (!excludeTable.contains(extractContent)) {
                        deleteTables.add(table);
                    } else {
                        int index = document.getPosOfTable(table);
                        keepParagraphs.add(index);
                        System.out.println("keeps:" + index);
                    }
                }
            }

            deleteTables.stream().forEach(item -> {
                int index = document.getPosOfTable(item);
                boolean success = document.removeBodyElement(index);
                System.out.println("tables:" + index + "--delete:" +success);
            });

//            testXWPFParagraph(document, keepParagraphs);

            deleteParagraphs(document, 6, document.getTables().size());

            FileOutputStream out = new FileOutputStream("./reports/newdemo.docx");
            document.write(out);

        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public void testXWPFParagraph(XWPFDocument document, Set<Integer> keeps) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        final int start = 9;
        final int end = paragraphs.size();
        int count = 0;
        Set<XWPFParagraph> deleteParagraphs = new HashSet<>();
        for (XWPFParagraph paragraph : paragraphs) {
            if (count > end) {
                break;
            }

            if (count > start && count <= end) {
                deleteParagraphs.add(paragraph);
            }
            count++;
        }

        deleteParagraphs.forEach(item -> {
            int index = document.getPosOfParagraph(item);
            System.out.println();

            if (!keeps.contains(index)) {
                boolean success = document.removeBodyElement(index);
                System.out.println("paragraph:" + index + "--delete:" + success);

            }
        });

    }

    public void deleteParagraphs(XWPFDocument document, int start, int end) {
        Iterator<IBodyElement> iterators = document.getBodyElementsIterator();
        int tableCount = 0;
        boolean canRecord = false;
        Set<XWPFParagraph> deleteParagraphs = new HashSet<>();

        while (iterators.hasNext()) {
            IBodyElement element = iterators.next();
            switch (element.getElementType()) {
                case TABLE:
                    tableCount++;
                    if (canRecord && tableCount > start) {
                        if (tableCount != end) {
                            // 最后一个试验table不需要保留paragraph
                            iterators.next();
                        }
                    }

                    if (!canRecord && tableCount >= start) {
                        canRecord = true;
                        // 每个试验table后续跟了一个paragraph,必须保留
                        iterators.next();
                    }
                    break;
                case PARAGRAPH:
                    if (canRecord) {
                        deleteParagraphs.add((XWPFParagraph) element);
                    }
                    break;
            }
        }

        deleteParagraphs.forEach(item -> {
            int index = document.getPosOfParagraph(item);
            boolean success = document.removeBodyElement(index);
            System.out.println("paragraph:" + index + "--" + success);

        });

    }
}
