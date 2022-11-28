package com.ashideng.report.template;

import org.apache.poi.xwpf.usermodel.*;

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
            List<XWPFTableRow>  rows = table.getRows();
            AtomicInteger index = new AtomicInteger();
            rows.forEach(item -> {
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

            doc.write(out);
        } catch (IOException e) {
            System.out.println(e);
            return false;
        }

        return true;
    }
}
