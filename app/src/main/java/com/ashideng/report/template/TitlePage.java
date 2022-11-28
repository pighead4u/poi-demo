package com.ashideng.report.template;

import org.apache.poi.xwpf.usermodel.*;

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

    public boolean createTableByRow(String path) {
        try (XWPFDocument doc = new XWPFDocument(); FileOutputStream out = new FileOutputStream(path)) {

        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        return true;
    }

//    public void test() {
//        XWPFTable Table = document.getTableArray(16);
//
//        XWPFTableRow getRow1 = Table.getRow(1);
//
//
//        XWPFTableRow getRow0 = Table.getRow(0);
//
//        //baris 1
//        for(int i = 0; i < listLHA.size(); i++) {
//            getRow0.getCell(0).setText(listLHA.get(0).getKeyProsses()+ " KEY PROSES");
//            break;
//        }
//
//        //baris 2
//
//        for(int i = 0; i < listLHA.size(); i++) {
//            getRow1.getCell(0).setText(listLHA.get(0).getRiskRating());
//            getRow1.getCell(1).setText(listLHA.get(0).getAuditObservationTitle()+ " AO TITLE");
//            break;
//        }
//
//        XWPFTableRow examInfoRow = Table.createRow();
//        XWPFTableCell cellRowInfo = examInfoRow.addNewTableCell();
//
//        XWPFParagraph examInfoRowP = cellRowInfo.getParagraphs().get(0);
//        XWPFRun examRun = examInfoRowP.createRun(); //problem 1
//
//        examInfoRowP.setAlignment(ParagraphAlignment.LEFT);
//        //list Action plan
//        examRun.setText("Action Plan:");
//        examRun.addBreak();
//        for (AuditEngagementLHA lha : listLHA) {
//            int i = listLHA.indexOf(lha);
//            examRun.setText(i+1 +"."+lha.getDescAP().replaceAll("\\<[^>]*>",""));
//            examRun.addBreak();
//        }
//        for(int i = 0; i < listLHA.size(); i++) {
//            examRun.setText("Target Date: ");
//            examRun.setText(listLHA.get(0).getTargetDateAP());
//            examRun.addBreak();
//            break;
//        }
//
//        examRun.addBreak();
//        for(int i = 0; i < listLHA.size(); i++) {
//            examInfoRow.getCell(0).setText(listLHA.get(0).getDescAO()+" Desc AO");
//            examRun.addBreak();
//            break;
//        }
//        //List penanggung jawab
//        examRun.setText("Penanggung Jawab:");
//        examRun.addBreak();
//        for (AuditEngagementLHA lha : listLHA) {
//            int i = listLHA.indexOf(lha);
//
//            examRun.setText(i+1 +"."+lha.getPicAP()+" - ");
//            examRun.setText(lha.getJabatanPicAP());
//            examRun.addBreak();
//        }
//    }
}
