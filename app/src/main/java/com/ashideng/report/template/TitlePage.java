package com.ashideng.report.template;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @Description: 创建封面
 * @Author: zhuhuanhuan@shunjiantech.cn
 * @Date: 2022/11/28 上午10:22
 * @Version: 1.0.0
 **/
public class TitlePage {
    List<Integer> HEIGHTS = Arrays.asList(3625, 93, 828, 1105, 386, 1018);
    List<String> CONTENTS = Arrays.asList("", "检测报告", "Test Report", "编号(No.)：{{ report.num }}", "", "");

    public boolean createPage(String path) {
        try (XWPFDocument doc = new XWPFDocument(); FileOutputStream out = new FileOutputStream(path)) {
            XWPFTable table = doc.createTable(6, 1);
            table.setWidth(8440);
            List<XWPFTableRow>  rows = table.getRows();
            AtomicInteger index = new AtomicInteger();
            rows.forEach(item -> {
//                if (index.get() == 0) {
//                    item.addNewTableCell();
//                    item.addNewTableCell();
//                }

                item.setHeight(HEIGHTS.get(index.get()));
                List<XWPFTableCell> cells = item.getTableCells();
//                cells.get(0).setText(CONTENTS.get(index.get()));
                XWPFRun run = cells.get(0).getParagraphs().get(0).createRun();
                run.setText(CONTENTS.get(index.get()));
                run.setFontSize(17);
                // 垂直剧中
                cells.get(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                // 水平居中
                cells.get(0).getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);

                index.getAndIncrement();
            });

            XWPFTable table2 = doc.createTable(6, 2);
            table2.setWidth(4220);

            doc.write(out);
        } catch (IOException e) {
            System.out.println(e);
            return false;
        }

        return true;
    }

    public boolean createTableByRow(String path, int rows, int columns) {
        try (XWPFDocument doc = new XWPFDocument(); FileOutputStream out = new FileOutputStream(path)) {
            XWPFTable table = doc.createTable(rows, columns);
            table.setWidth("100%");
            table.setWidthType(TableWidthType.PCT);//设置表格相对宽度
            table.setTableAlignment(TableRowAlign.CENTER);

            //合并单元格
            XWPFTableRow row1 = table.getRow(0);
            mergeCells(row1, 0, 2);

            doc.write(out);
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return true;
    }

    static class TitlePagePO {

        private final String title_cn;
        private int title_cn_size;
        private int title_cn_height;
        private final String title_en;
        private int title_en_size;
        private int title_en_height;
        private final String report_num;
        private int report_num_size;

        public String getTitle_cn() {
            return title_cn;
        }

        public int getTitle_cn_size() {
            return title_cn_size;
        }

        public String getTitle_en() {
            return title_en;
        }

        public int getTitle_en_size() {
            return title_en_size;
        }

        public String getReport_num() {
            return report_num;
        }

        public int getReport_num_size() {
            return report_num_size;
        }

        public TitlePagePO(String ititle_cn,
                           int ititle_cn_size,
                           String ititle_en,
                           int ititle_en_size,
                           String ireport_num,
                           int ireport_num_size) {
            title_cn = ititle_cn;
            title_cn_size = ititle_cn_size;
            title_en = ititle_en;
            title_en_size = ititle_en_size;
            report_num = ireport_num;
            report_num_size = ireport_num_size;
        }
    }

    private boolean mergeCells(XWPFTableRow row, int start_index, int end_index) {
        for (int i = start_index; i <= end_index; i++) {
            if (i == start_index) {
                XWPFTableCell cell1 = row.getCell(start_index);
                CTTcPr cellCtPr = getCellCTTcPr(cell1);
                cellCtPr.addNewHMerge().setVal(STMerge.RESTART);
            } else {
                XWPFTableCell cell2 = row.getCell(i);
                CTTcPr cellCtPr2 = getCellCTTcPr(cell2);
                cellCtPr2.addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }

        return true;
    }

    private CTTcPr getCellCTTcPr(XWPFTableCell cell) {
        CTTc cttc = cell.getCTTc();
        CTTcPr tcPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();
        return tcPr;
    }

}
