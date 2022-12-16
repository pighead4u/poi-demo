package com.ashideng.report.template;

import org.apache.commons.lang3.time.StopWatch;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * @author zhuhuanhuan@shunjiantech.cn
 * @version 1.0.0
 * @description: 正式的模板，方便后续的代码填充占位符
 * @create 2022/12/9 下午2:34
 **/
public class FormalTemplate {

    // 定位试验table的正则
    private final String PATTERN = "(?<=\\{\\{\\^)[^\\}\\}]+";
    private final int TABLE_START = 6;

    private  final String TABLE_NAME = "2检测结果汇总";

    public void buildNewTemplate(String path, Set<String> excludeTable) {
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(path))) {
            StopWatch watch = new StopWatch();
            watch.start();

            deleteExperiments(document, excludeTable);
            watch.split();
            System.out.println("deleteExperiments-time:" + watch.getTime());

            deleteResults(document, excludeTable);
            watch.split();
            System.out.println("deleteResults-time:" + watch.getTime());

            deleteTables(excludeTable, document);
            watch.split();
            System.out.println("deleteTables-time:" + watch.getTime());

            deleteParagraphs(document, TABLE_START, document.getTables().size());
            watch.split();
            System.out.println("deleteParagraphs-time:" + watch.getTime());

            FileOutputStream out = new FileOutputStream("./reports/newdemo.docx");
            document.write(out);

            watch.stop();
            System.out.println("total_time-time:" + watch.getTime());


        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private void deleteExperiments(XWPFDocument document, Set<String> excludeTable) {
        List<XWPFTable> tables = document.getTables();
        XWPFTable table = tables.get(3);
        List<XWPFTableRow> rows = table.getRows();

        for (int i = 0; i < rows.size(); i++) {
            XWPFTableRow row = rows.get(i);
//            System.out.println("rows:" + i);
            boolean isBreak = true;
            for (int j = 0; j < row.getTableCells().size(); j++) {
                XWPFTableCell cell = row.getTableCells().get(j);
                String cellContent = cell.getText();
                // 定位到检测项目这一行
                if (cellContent.contains("Test item")) {
                    isBreak = false;
                    continue;
                }

                if (isBreak) {
                    break;
                }

                // 总体思路：一行一行遍历，如果在保留的试验项目内的，这一行保留，并删除后续的定位符号;如果不在，则删除这一行
                List<Integer> deleteParagraphs = new ArrayList<>();
                for (int k = 0; k < cell.getParagraphs().size(); k++) {
                    XWPFParagraph paragraph = cell.getParagraphArray(k);
                    String content = paragraph.getText();
                    boolean isExisted = false;
                    for (String item : excludeTable) {
                        if (content.contains(item)) {
                            isExisted = true;

                            List<XWPFRun> runs = paragraph.getRuns();
                            boolean isNeedDelete = true;
                            for (int m = runs.size()-1; m > 0;m--) {
                                XWPFRun run = paragraph.getRuns().get(m);
                                String deleteContent = run.text();
                                if (!isNeedDelete && deleteContent.contains("}}")) {
                                    break;
                                }

                                // }}{{-是在一个run里的
                                if (deleteContent.contains("{{")) {
                                    isNeedDelete = false;
                                    run.setText("}}", 0);
                                }

                                if (isNeedDelete) {
                                    paragraph.removeRun(m);
                                }
                            }

                            break;
                        }
                    }

                    if (!isExisted) {
                        deleteParagraphs.add(k);
                    }

                }

                for (int k = deleteParagraphs.size()-1; k > 0; k--) {
                    cell.removeParagraph(deleteParagraphs.get(k));
                }

            }

        }
    }

    private void deleteResults(XWPFDocument document, Set<String> excludeTable) {
        List<XWPFTable> tables = document.getTables();
        // get the table number
        int tableNum = 0;
        for (int i = 0; i < tables.size(); i++) {
            if (tables.get(i).getRow(0).getCell(0).getText().equals(TABLE_NAME)) {
                tableNum = i;
                break;
            }
        }

        XWPFTable table = tables.get(tableNum); // the table 2检测结果汇总
        List<XWPFTableRow> rows = table.getRows();

        // create the pattern for regex matching
        String p = "([a-z]+_.+)\\.";
        Pattern pattern = Pattern.compile(p);

        // store all possible outcomes
        List<Integer> rowIndices = IntStream.range(2, rows.size()).boxed().collect(Collectors.toList());

        // save wanted lines
        for (int i = 2; i < rows.size(); i++) { // since the content starts at the third line
            for (XWPFTableCell cell : rows.get(i).getTableCells()) {
                Matcher matcher = pattern.matcher(cell.getText());
                if (matcher.find() && excludeTable.contains(matcher.group(1))) {
                    rowIndices.remove((Integer) i);
                    break;
                }
            }
        }

        System.out.println(rowIndices);
        Collections.reverse(rowIndices);

        // keep only the rows that match the pattern
        for (int i : rowIndices) {
            table.removeRow(i);
        }
    }

    private void deleteTables(Set<String> excludeTable, XWPFDocument document) {
        List<XWPFTable> tables = document.getTables();

        Set<XWPFTable> deleteTables = new HashSet<>();
        Pattern pattern = Pattern.compile(PATTERN);

        for (int i = TABLE_START; i < tables.size(); i++) {
            XWPFTable table = tables.get(i);
            String content = table.getText();
            Matcher matcher = pattern.matcher(content);
            if (matcher.find()) {
                String extractContent = matcher.group();
                if (!excludeTable.contains(extractContent)) {
                    deleteTables.add(table);
                }
            }
        }

        deleteTables.stream().forEach(item -> {
            int index = document.getPosOfTable(item);
            boolean success = document.removeBodyElement(index);
            System.out.println("tables:" + index + "--delete:" + success);
        });
    }

    private void testXWPFParagraph(XWPFDocument document, Set<Integer> keeps) {
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
