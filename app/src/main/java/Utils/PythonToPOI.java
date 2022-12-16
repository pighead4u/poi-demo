package Utils;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PythonToPOI {
    private static final double FONT_SIZE = 10.5;
    private static final String FONT_FAMILY = "Times New Roman";

    public static int findFirstValidCell(XWPFTableRow row) {
        int res = 0;
        for (int i = 0; i < row.getTableCells().size() - 1; i++) {
            if (row.getTableCells().get(i) != null) {
                res = i;
                break;
            }
        }
        return res;
    }

    public static int findFirstValidParagraph(XWPFTableCell cell) {
        int res = 0;
        for (int i = 0; i < cell.getParagraphs().size() - 1; i++) {
            if (cell.getParagraphs().get(i) != null) {
                res = i;
                break;
            }
        }
        return res;
    }
    public static int findFirstValidRun(XWPFParagraph para) {
        int res = 0;
        for (int i = 0; i < para.getRuns().size() - 1; i++) {
            if (para.getRuns().get(i) != null) {
                res = i;
                break;
            }
        }
        return res;
    }

    public static void clearRuns(XWPFParagraph targetParagraph) {
        if (!targetParagraph.isEmpty()) {
            for (int i = targetParagraph.getRuns().size() - 1; i >= 0; i--) {
                targetParagraph.removeRun(i);
            }
        }
    }

    public static void clearRow(XWPFTableRow targetRow) {
        if (targetRow.getTableCells().size() != 0) {
            for (int i = targetRow.getTableCells().size() - 1; i >= 0; i--) {
                XWPFTableCell cell = targetRow.getTableCells().get(i);
                if (cell.getParagraphs().size() != 0) {
                    for (int j = cell.getParagraphs().size() - 1; j >= 0; j--) {
                        XWPFParagraph para = cell.getParagraphs().get(j);
                        if (para.getRuns().size() != 0) {
                            for (int k = para.getRuns().size(); k >= 0; k--) {
                                para.removeRun(k);
                            }
                            cell.removeParagraph(j);
                        }
                    }
                }
                targetRow.removeCell(i);
            }
        }
    }

    public static void setTextStyleInRun(XWPFRun run) {
        run.setFontSize(FONT_SIZE);
        run.setFontFamily(FONT_FAMILY);
    }

    public static void addContentToRow(XWPFTableRow targetRow, String content) {
        XWPFRun newRun = targetRow.createCell().addParagraph().createRun();
        newRun.setText(content);
        newRun.setFontSize(FONT_SIZE);
        newRun.setFontFamily(FONT_FAMILY);


    }

    public static void insertRow(XWPFTable table, int copyrowIndex, int newrowIndex) {
        // 在表格中指定的位置新增一行
        XWPFTableRow targetRow = table.insertNewTableRow(newrowIndex);
        // 获取需要复制行对象
        XWPFTableRow copyRow = table.getRow(copyrowIndex);
        //复制行对象
        targetRow.getCtRow().setTrPr(copyRow.getCtRow().getTrPr());
        //或许需要复制的行的列
        List<XWPFTableCell> copyCells = copyRow.getTableCells();
        //复制列对象
        XWPFTableCell targetCell = null;
        for (int i = 0; i < copyCells.size(); i++) {
            XWPFTableCell copyCell = copyCells.get(i);
            targetCell = targetRow.addNewTableCell();
            targetCell.getCTTc().setTcPr(copyCell.getCTTc().getTcPr());
            if (copyCell.getParagraphs() != null && copyCell.getParagraphs().size() > 0) {
                targetCell.getParagraphs().get(0).getCTP().setPPr(copyCell.getParagraphs().get(0).getCTP().getPPr());
                if (copyCell.getParagraphs().get(0).getRuns() != null
                        && copyCell.getParagraphs().get(0).getRuns().size() > 0) {
                    XWPFRun cellR = targetCell.getParagraphs().get(0).createRun();
                    cellR.setBold(copyCell.getParagraphs().get(0).getRuns().get(0).isBold());
                }
            }
        }

    }

    private static void createCellsAndCopyStyles(XWPFTableRow targetRow, XWPFTableRow sourceRow) {
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        List<XWPFTableCell> tableCells = sourceRow.getTableCells();
        if (CollectionUtils.isEmpty(tableCells)) {
            return;
        }
        for (XWPFTableCell sourceCell : tableCells) {
            XWPFTableCell newCell = targetRow.addNewTableCell();
            newCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            List<XWPFParagraph> sourceParagraphs = sourceCell.getParagraphs();
            if (CollectionUtils.isEmpty(sourceParagraphs)) {
                continue;
            }
            XWPFParagraph sourceParagraph = sourceParagraphs.get(0);
            List<XWPFParagraph> targetParagraphs = newCell.getParagraphs();
            if (CollectionUtils.isEmpty(targetParagraphs)) {
                XWPFParagraph p = newCell.addParagraph();
                p.getCTP().setPPr(sourceParagraph.getCTP().getPPr());
                XWPFRun run = p.getRuns().isEmpty() ? p.createRun() : p.getRuns().get(0);
                run.setFontFamily(sourceParagraph.getRuns().get(0).getFontFamily());
            } else {
                XWPFParagraph p = targetParagraphs.get(0);
                p.getCTP().setPPr(sourceParagraph.getCTP().getPPr());
                XWPFRun run = p.getRuns().isEmpty() ? p.createRun() : p.getRuns().get(0);
                run.setFontFamily(sourceParagraph.getRuns().get(0).getFontFamily());
            }
        }
    }


    public static void handles(String filePath) {
        try (XWPFDocument tempDoc = new XWPFDocument(new FileInputStream(filePath))) {


            List<XWPFTable> tables = tempDoc.getTables();
            // {%tr for item in dts_kgg_jydzsy %}
            String loopRegex = "%.*for.?(.*)in\\s((.*)\\.|.*).*\\s%"; // {%tr for instrument in hwg_gtczjhdcl.instrument %}
            String nationalStdRegex = "%.*for\\s(.*)\\sin\\s(.*)\\s%}.*endfor.*";
            String endloopRegex = "%.*(endfor).*%"; // {%tr endfor %}
            String multiFieldsInOneCellRegex = "\\{\\{(@)?(.*)\\.(.*)}}";
            Pattern loopPattern = Pattern.compile(loopRegex);
            Pattern nationalStdPattern = Pattern.compile(nationalStdRegex);
            Pattern endloopPattern = Pattern.compile(endloopRegex);
            Pattern multiFieldsInOneCellPattern = Pattern.compile(multiFieldsInOneCellRegex);

            int tabAcc = 0;
            for (XWPFTable table : tables) {
                System.out.println("table number       " + tabAcc);
                System.out.println();
                tabAcc += 1;
                List<Integer> rowsToDelete = new ArrayList<>();
                for (int i = 0; i < table.getRows().size(); i++) {
                    XWPFTableRow row = table.getRow(i);
                    for (int j = 0; j < row.getTableCells().size(); j++) {
                        String cellText = row.getCell(j).getText();
                        Matcher loopMatcher = loopPattern.matcher(cellText); // match the Python-style for loop
                        Matcher nationalStdMatcher = nationalStdPattern.matcher(cellText);
                        Matcher endloopMatcher = endloopPattern.matcher(cellText);

                        /* If it's the case 检测依据 e.g., section 1.2 */
                        if (nationalStdMatcher.find()) {
                            XWPFTableCell cellToPlace = table.getRow(i - 1).getCell(0);

                            XWPFRun cellToPlaceNewRun = cellToPlace.getParagraphs().get(cellToPlace.getParagraphs().size() - 1).createRun();
                            cellToPlaceNewRun.setText("{" + nationalStdMatcher.group(2) + "}");
                            cellToPlaceNewRun.setFontSize(FONT_SIZE);
                            cellToPlaceNewRun.setFontFamily(FONT_FAMILY);

                            XWPFTableCell currCell = table.getRow(i).getCell(0);


                            for (int r = currCell.getParagraphs().size() - 1; r >= 0; r--) {
                                currCell.removeParagraph(r);
                            }
                            XWPFRun currCellNewRun = currCell.addParagraph().createRun();
                            currCellNewRun.setText("[" + nationalStdMatcher.group(1) + "]");
                            currCellNewRun.setFontSize(FONT_SIZE);
                            currCellNewRun.setFontFamily(FONT_FAMILY);

                        }
                        else if (loopMatcher.find()) { // deals with the case if the cell contains the Python-style for loop

                            // REASONING: There are two sorts of content to place in the row above the for loop:
                            // one with a dot like "hwg_gtczjhdcl.instrument", and the other without a dot like
                            // "dts_kgg_jydzsy". We only want the tag so we can iterate the items in a POI-tl api
                            // fashion, so when "." is present, we'd take the tag before it.
                            String loopContentToPlace = !loopMatcher.group(2).contains(".") ? loopMatcher.group(2) :
                                    loopMatcher.group(3);

                            /* check if there exists previous row */
                            XWPFTableRow prevRow;

                            // TODO: FIX ADDING A NEW ROW ISSUE FOR demo.docx AT PAGE 41

                            if (i == 0) {
//                                insertRow(table, i, i);
//                                table.getRow(0).createCell().addParagraph().createRun().setText("aaaaaa");

                                System.out.println("cells   " + table.getRow(0).getTableCells().size());
                                System.out.println("paras   " + table.getRow(0).getTableCells().get(0).getParagraphs().size());
                                System.out.println("runs   " + table.getRow(0).getTableCells().get(0).getParagraphs().get(0).getRuns().size());
                                clearRuns(table.getRow(i).getTableCells().get(0).getParagraphs().get(0));
                                table.getRow(i).getTableCells().get(0).getParagraphs().get(0).createRun().setText("{" + loopContentToPlace + "}");
                                setTextStyleInRun(table.getRow(i).getTableCells().get(0).getParagraphs().get(0).getRuns().get(0));



                                break;
                            }
                            else {
                                prevRow = table.getRow(i - 1);
                            }


                            XWPFTableCell cellToPlace = prevRow.getCell(0);

                            XWPFRun newRun = cellToPlace.getParagraphs().get(cellToPlace.getParagraphs().size() - 1).
                                    createRun();
                            newRun.setText("{" + loopContentToPlace + "}");
                            newRun.setFontFamily("Times New Roman");
                            newRun.setFontSize(10.5);


                            rowsToDelete.add(i); // saves the rows that contain "for ... in ..."
                            XWPFTableRow nextRow = table.getRow(i + 1);

                            // get each cell from the next row
                            for (XWPFTableCell cell : nextRow.getTableCells()) {
                                String nextRowCellText = cell.getText();
                                Matcher multiFieldsInOneCellMatcher = multiFieldsInOneCellPattern.
                                        matcher(nextRowCellText);

                                if (multiFieldsInOneCellMatcher.find()) {

                                    String multiFieldsMatchGroup_1 = multiFieldsInOneCellMatcher.group(1);
                                    String cellMatcherResult = (multiFieldsMatchGroup_1 != null
                                            && multiFieldsMatchGroup_1.equals("@"))
                                            ? multiFieldsMatchGroup_1 + multiFieldsInOneCellMatcher.group(3)
                                            : multiFieldsInOneCellMatcher.group(3);


                                    XWPFParagraph paragraph = cell.getParagraphs().get(findFirstValidParagraph(cell));

                                    // remove all old cells
                                    ParagraphAlignment paragraphAlignment = paragraph.getAlignment();

                                    for (int r = cell.getParagraphs().size() - 1; r >= 0; r--) {
                                        cell.removeParagraph(r);
                                    }

                                    // add new content to cells
                                    XWPFParagraph newParagraph = cell.addParagraph();
                                    newParagraph.createRun().setText("[" + cellMatcherResult + "]");
                                    newParagraph.getRuns().get(0).setFontSize(FONT_SIZE);
                                    newParagraph.getRuns().get(0).setFontFamily(FONT_FAMILY);
                                    newParagraph.setAlignment(paragraphAlignment);
                                }
                            }
                        }

                        else if (endloopMatcher.find()) {
                            rowsToDelete.add(i); // saves the rows that contain "endloop"
                        }
                    }
                }
//                System.out.println("rows to delete: ");
//                System.out.println(rowsToDelete);

                for (int i = rowsToDelete.size() - 1; i >= 0; i--) {
                    table.removeRow(rowsToDelete.get(i));
                }

            }

            FileOutputStream fos = new FileOutputStream("./output/output.docx");
            tempDoc.write(fos);
        } catch (IOException e) {
            System.out.println("Bad File or Bad File Path");
        }
    }

    private static void foo() {
        String loopRegex = "%.*for.*in.(.*)\\..*%"; // {%tr for instrument in hwg_gtczjhdcl.instrument %}
        String endloopRegex = "%.*(endfor).*%"; // {%tr endfor %}
        String multiFieldsInOneCellRegex = "(@?)(?:.*\\.)(.*)}}";

        String s1 = "{{@dts_kgg_jdldcjsy_item.picpath}}";
        Pattern p1 = Pattern.compile(multiFieldsInOneCellRegex);
        Matcher m1 = p1.matcher(s1);

    }

    public static void main(String[] args) {
        String path = "./templates/demo.docx";

        handles(path);
    }
}
