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

    /**
     * Finds the index of the first valid cell in a row
     * @param row the row to traverse
     * @return the index of the first cell; -1 if fails to find
     */
    public static int findFirstValidCell(XWPFTableRow row) {
        int res = -1;
        for (int i = 0; i < row.getTableCells().size() - 1; i++) {
            if (row.getTableCells().get(i) != null) {
                return i;
            }
        }
        return res;
    }

    /**
     * Finds the index of the first valid paragraph in a cell
     * @param cell the cell to traverse
     * @return the index of the first paragraph; -1 if fails to find
     */
    public static int findFirstValidParagraph(XWPFTableCell cell) {
        int res = -1;
        for (int i = 0; i < cell.getParagraphs().size() - 1; i++) {
            if (cell.getParagraphs().get(i) != null) {
                return i;
            }
        }
        return res;
    }

    /**
     * Finds the index of the first valid run in a paragraph
     * @param para the paragraph to traverse
     * @return the index of the first run; -1 if fails to find
     */
    public static int findFirstValidRun(XWPFParagraph para) {
        int res = -1;
        for (int i = 0; i < para.getRuns().size() - 1; i++) {
            if (para.getRuns().get(i) != null) {
                return i;
            }
        }
        return res;
    }

    /**
     * Clears all runs in a specific paragraph
     * @param targetParagraph the paragraph to clear with
     */
    public static void clearRuns(XWPFParagraph targetParagraph) {
        if (!targetParagraph.isEmpty()) {
            for (int i = targetParagraph.getRuns().size() - 1; i >= 0; i--) {
                targetParagraph.removeRun(i);
            }
        }
    }

    /**
     * Deletes all elements in a row
     * @param targetRow the row in which to delete
     */
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

    /**
     * Converts all Python-style elements like for loops, endloop marks, picture fields start with "@", etc. into
     * POI-tl style elements.
     * @param inputDir the directory where files to handle with are at
     * @param outputDir the directory to output the processed files
     * @param fileName the name of the file to handle with
     */
    public static void handles(String inputDir, String outputDir, String fileName) {
        try (XWPFDocument tempDoc = new XWPFDocument(new FileInputStream(inputDir + fileName))) {
            List<XWPFTable> tables = tempDoc.getTables();

            for (XWPFTable table : tables) {
                List<Integer> rowsToDelete = new ArrayList<>();
                for (int i = 0; i < table.getRows().size(); i++) {
                    XWPFTableRow row = table.getRow(i);
                    for (int j = 0; j < row.getTableCells().size(); j++) {
                        String cellText = row.getCell(j).getText();

                        /* generate corresponding matchers */
                        Matcher loopMatcher = LOOP_PATTERN.matcher(cellText); // match the Python-style for loop
                        Matcher nationalStdMatcher = NATIONAL_STD_PATTERN.matcher(cellText); // match a
                        Matcher endloopMatcher = ENDLOOP_PATTERN.matcher(cellText);

                        /* If it's the case 检测依据 e.g., see section 1.2 */
                        if (nationalStdMatcher.find()) {
                            XWPFTableCell cellToPlace = table.getRow(i - 1).getCell(0);

                            XWPFRun cellToPlaceNewRun = cellToPlace.getParagraphs().
                                    get(cellToPlace.getParagraphs().size() - 1).createRun();
                            cellToPlaceNewRun.setText("{" + nationalStdMatcher.group(2) + "}");
                            setTextStyleInRun(cellToPlaceNewRun);
                            XWPFTableCell currCell = table.getRow(i).getCell(0);

                            for (int r = currCell.getParagraphs().size() - 1; r >= 0; r--) {
                                currCell.removeParagraph(r);
                            }
                            XWPFRun currCellNewRun = currCell.addParagraph().createRun();
                            currCellNewRun.setText("[" + nationalStdMatcher.group(1) + "]");
                            setTextStyleInRun(currCellNewRun);
                        }
                        else if (loopMatcher.find()) { // deals with the case the cell contains a Python-style for loop

                            /* 针对例如"hwg_gtczjhdcl.instrument"与"dts_kgg_jydzsy"两种情况的处理 (有无 ".") */
                            String loopContentToPlace = !loopMatcher.group(2).contains(".") ? loopMatcher.group(2) :
                                    loopMatcher.group(3);

                            /* 针对例如 demo.docx 2.22 ~ 2.23 之间的表的情况处理 (匹配到的行为第0行)
                             * 此情况下, 删去原row内容后直接添加新内容*/
                            if (i == 0) {
                                // TODO: FIX THE ISSUE WHERE THE ROWS AFTER THIS ONE CANNOT BE EDITED
                                XWPFParagraph paragraph = table.getRow(i).getTableCells().get(0).getParagraphs().get(0);
                                clearRuns(paragraph);
                                paragraph.createRun().setText("{" + loopContentToPlace + "}");
                                setTextStyleInRun(paragraph.getRuns().get(0));
                            }
                            else {
                                XWPFTableRow prevRow = table.getRow(i - 1);
                                rowsToDelete.add(i);
                                XWPFTableCell cellToPlace = prevRow.getCell(0);

                                XWPFRun newRun = cellToPlace.getParagraphs().
                                        get(cellToPlace.getParagraphs().size() - 1).createRun();
                                newRun.setText("{" + loopContentToPlace + "}");
                                setTextStyleInRun(newRun);
                            }

                            XWPFTableRow nextRow = table.getRow(i + 1);

                            /* 通常情况, 将后一行的cell中的原有的字段转化为POI样式的字段 */
                            for (XWPFTableCell cell : nextRow.getTableCells()) {
                                String nextRowCellText = cell.getText();
                                Matcher multiFieldsMatcher = MULTI_FIELDS_PATTERN.
                                        matcher(nextRowCellText);

                                if (multiFieldsMatcher.find()) { // 处理表格中一个格子里有一个或多个cell的情况
                                    String multiFieldsMatchGroup_1 = multiFieldsMatcher.group(1);
                                    String cellMatcherResult = (multiFieldsMatchGroup_1 != null
                                            && multiFieldsMatchGroup_1.equals("@")) // 若有@, 说明该字段为图片且需特殊处理
                                            ? multiFieldsMatchGroup_1 + multiFieldsMatcher.group(3)
                                            : multiFieldsMatcher.group(3);

                                    XWPFParagraph paragraph = findFirstValidParagraph(cell) != -1
                                            ? cell.getParagraphs().get(findFirstValidParagraph(cell))
                                            : cell.addParagraph();
                                    ParagraphAlignment paragraphAlignment = paragraph.getAlignment(); // 储存原有格式

                                    for (int r = cell.getParagraphs().size() - 1; r >= 0; r--) {
                                        cell.removeParagraph(r);
                                    }

                                    /* 给每个cell添加样式为 [fieldName] 的字段 */
                                    XWPFParagraph newParagraph = cell.addParagraph();
                                    XWPFRun run = newParagraph.createRun();
                                    run.setText("[" + cellMatcherResult + "]");
                                    setTextStyleInRun(run);
                                    newParagraph.setAlignment(paragraphAlignment);
                                }
                            }
                        }
                        /* 直接删去row, 若中含有"endloop" */
                        else if (endloopMatcher.find()) {
                            rowsToDelete.add(i); // saves the rows that contain "endloop"
                        }
                    }
                }
                for (int i = rowsToDelete.size() - 1; i >= 0; i--) {
                    table.removeRow(rowsToDelete.get(i));
                }
            }
            /* 将doc写入对应的文件中 */
            FileOutputStream fos = new FileOutputStream(outputDir + "output_" + fileName);
            tempDoc.write(fos);
        } catch (IOException e) {
            System.out.println("Bad File or Bad File Path");
        }
    }

    public static void main(String[] args) {
        File file = new File("./templates/");
        File[] files = file.listFiles();
        for (File f : files) {
            System.out.println("./output/"+f.getName());
            handles("./templates/", "./output/", f.getName());
        }
//        String path = "./templates/demo.docx";
//        handles(path);
    }
}
