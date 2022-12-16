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

    public static final String LOOP_REGEX = "%.*for.?(.*)in\\s((.*)\\.|.*).*\\s%"; // {%tr for instrument in hwg_gtczjhdcl.instrument %}
    public static final String NATIONAL_STD_REGEX = "%.*for\\s(.*)\\sin\\s(.*)\\s%}.*endfor.*";
    public static final String ENDLOOP_REGEX = "%.*(endfor).*%"; // {%tr endfor %}
    public static final String MULTI_FIELDS_REGEX = "\\{\\{(@)?(.*)\\.(.*)}}";

    public static Pattern LOOP_PATTERN = Pattern.compile(LOOP_REGEX);
    public static Pattern NATIONAL_STD_PATTERN = Pattern.compile(NATIONAL_STD_REGEX);
    public static Pattern ENDLOOP_PATTERN = Pattern.compile(ENDLOOP_REGEX);
    public static Pattern MULTI_FIELDS_PATTERN = Pattern.compile(MULTI_FIELDS_REGEX);

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
        setTextStyleInRun(newRun);
    }


    public static void handles(String filePath) {
        try (XWPFDocument tempDoc = new XWPFDocument(new FileInputStream(filePath))) {
            List<XWPFTable> tables = tempDoc.getTables();
            // {%tr for item in dts_kgg_jydzsy %}

            int tabAcc = 0;
            for (XWPFTable table : tables) {
                tabAcc += 1;
                List<Integer> rowsToDelete = new ArrayList<>();
                for (int i = 0; i < table.getRows().size(); i++) {
                    XWPFTableRow row = table.getRow(i);
                    for (int j = 0; j < row.getTableCells().size(); j++) {
                        String cellText = row.getCell(j).getText();
                        Matcher loopMatcher = LOOP_PATTERN.matcher(cellText); // match the Python-style for loop
                        Matcher nationalStdMatcher = NATIONAL_STD_PATTERN.matcher(cellText);
                        Matcher endloopMatcher = ENDLOOP_PATTERN.matcher(cellText);

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
                            setTextStyleInRun(currCellNewRun);
                        }
                        else if (loopMatcher.find()) { // deals with the case if the cell contains the Python-style for loop

                            /* 针对例如"hwg_gtczjhdcl.instrument"与"dts_kgg_jydzsy"两种情况的处理 (有无 ".") */
                            String loopContentToPlace = !loopMatcher.group(2).contains(".") ? loopMatcher.group(2) :
                                    loopMatcher.group(3);


                            /* 针对例如 demo.docx 2.22 ~ 2.23 之间的表的情况处理 (匹配到的行为第0行) */
                            if (i == 0) {
                                clearRuns(table.getRow(i).getTableCells().get(0).getParagraphs().get(0));
                                table.getRow(i).getTableCells().get(0).getParagraphs().get(0).createRun().setText("{" + loopContentToPlace + "}");
                                setTextStyleInRun(table.getRow(i).getTableCells().get(0).getParagraphs().get(0).getRuns().get(0));
//                                break;
                            }
                            else {
                                XWPFTableRow prevRow = table.getRow(i - 1);
                                rowsToDelete.add(i);
                                XWPFTableCell cellToPlace = prevRow.getCell(0);

                                XWPFRun newRun = cellToPlace.getParagraphs().get(cellToPlace.getParagraphs().size() - 1).
                                        createRun();
                                newRun.setText("{" + loopContentToPlace + "}");
                                newRun.setFontFamily("Times New Roman");
                                newRun.setFontSize(10.5);
                            }

                            XWPFTableRow nextRow = table.getRow(i + 1);

                            // get each cell from the next row
                            for (XWPFTableCell cell : nextRow.getTableCells()) {
                                String nextRowCellText = cell.getText();
                                Matcher multiFieldsMatcher = MULTI_FIELDS_PATTERN.
                                        matcher(nextRowCellText);

                                if (multiFieldsMatcher.find()) { // 处理表格中一个格子里有一个或多个cell的情况
                                    String multiFieldsMatchGroup_1 = multiFieldsMatcher.group(1);
                                    String cellMatcherResult = (multiFieldsMatchGroup_1 != null
                                            && multiFieldsMatchGroup_1.equals("@"))
                                            ? multiFieldsMatchGroup_1 + multiFieldsMatcher.group(3)
                                            : multiFieldsMatcher.group(3);

                                    XWPFParagraph paragraph = cell.getParagraphs().get(findFirstValidParagraph(cell));

                                    ParagraphAlignment paragraphAlignment = paragraph.getAlignment();

                                    for (int r = cell.getParagraphs().size() - 1; r >= 0; r--) {
                                        cell.removeParagraph(r);
                                    }

                                    // add new content to cells
                                    XWPFParagraph newParagraph = cell.addParagraph();
                                    XWPFRun run = newParagraph.createRun();
                                    run.setText("[" + cellMatcherResult + "]");
                                    setTextStyleInRun(run);
                                    newParagraph.setAlignment(paragraphAlignment);
                                }
                            }
                        }

                        else if (endloopMatcher.find()) {
                            rowsToDelete.add(i); // saves the rows that contain "endloop"
                        }
                    }
                }
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


    public static void main(String[] args) {
        String path = "./templates/demo.docx";

        handles(path);
    }
}
