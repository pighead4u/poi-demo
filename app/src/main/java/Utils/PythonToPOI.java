package Utils;

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

    private static final String LOOP_REGEX = "%.*for.?(.*)in\\s((.*)\\.|.*).*%";
    private static final String NATIONAL_STD_REGEX = "%.*for\\s(.*)\\sin\\s(.*)\\s%}.*endfor.*";
    private static final String ENDLOOP_REGEX = "%.*(endfor).*%"; // {%tr endfor %}
    private static final String MULTI_FIELDS_REGEX = "\\{\\{(@)?(.*)\\.(.*)}}";

    private static final Pattern LOOP_PATTERN = Pattern.compile(LOOP_REGEX);
    private static final Pattern NATIONAL_STD_PATTERN = Pattern.compile(NATIONAL_STD_REGEX);
    private static final Pattern ENDLOOP_PATTERN = Pattern.compile(ENDLOOP_REGEX);
    private static final Pattern MULTI_FIELDS_PATTERN = Pattern.compile(MULTI_FIELDS_REGEX);

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

            int tabNum = 0;
            for (XWPFTable table : tables) {
                System.out.println(tabNum);
                tabNum ++;
                List<Integer> rowsToDelete = new ArrayList<>();
                for (int i = 0; i < table.getRows().size(); i++) {

                    System.out.println();
                    System.out.println();
                    System.out.println("row number:       " + i);

                    XWPFTableRow row = table.getRow(i);
                    for (int j = 0; j < row.getTableCells().size(); j++) {
                        String cellText = row.getCell(j).getText();
                        System.out.println("cell content:         " + cellText);
                        /* generate corresponding matchers */
                        Matcher loopMatcher = LOOP_PATTERN.matcher(cellText); // match the Python-style for loop
                        Matcher nationalStdMatcher = NATIONAL_STD_PATTERN.matcher(cellText); // match
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
                                XWPFParagraph paragraph = table.getRow(i).getTableCells().get(0).getParagraphs().get(0);
                                clearRuns(paragraph);
                                paragraph.createRun().setText("{" + loopContentToPlace + "}");
                                setTextStyleInRun(paragraph.getRuns().get(0));
                            }
                            else {
                                XWPFTableRow prevRow = table.getRow(i - 1);
                                rowsToDelete.add(i);

                                XWPFTableCell cellToPlace = prevRow.getCell(0);
                                if (cellToPlace.getText().trim().equals("")) {
                                    prevRow.removeCell(0);
                                    cellToPlace = prevRow.getCell(0);
                                }
                                System.out.println("-------- cellToPlace is empty?      " + cellToPlace.getText());

                                // TODO: EDIT TO FIX THE "CANNOT ADD CONTENT TO FIRST CELL" ISSUE
                                XWPFRun newRun = !cellToPlace.getParagraphs().isEmpty()
                                        ? cellToPlace.getParagraphs().get(cellToPlace.getParagraphs().size() - 1).createRun()
                                        : cellToPlace.addParagraph().createRun();
                                newRun.setText("{" + loopContentToPlace + "}");
                                System.out.println("test:        " + newRun);

                                setTextStyleInRun(newRun);
                            }

                            XWPFTableRow nextRow = table.getRow(i + 1);

                            /* 将后一行的cell中的原有的字段转化为POI样式的字段 */
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
                                    for (int r = cell.getParagraphs().size() - 1; r >= 0; r--) {
                                        cell.removeParagraph(r);
                                    }

                                    /* 给每个cell添加样式为 [fieldName] 的字段 */
                                    XWPFParagraph newParagraph = cell.addParagraph();
                                    XWPFRun run = newParagraph.createRun();
                                    run.setText("[" + cellMatcherResult + "]");
                                    setTextStyleInRun(run);
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
            tempDoc.close();
            fos.close();
        } catch (IOException e) {
            System.out.println("Bad File or Bad File Path");
        }
    }

    public static void main(String[] args) {
        File file = new File("./templates/");
        File[] files = file.listFiles();
        for (File f : files) {
            if (f.getName().contains("~$")) {
                continue;
            }
            System.out.println("./output/"+f.getName());
            if (f.getName().equals("高压开关柜.docx")) {
                handles("./templates/", "./output/", f.getName());
            }
        }
//        String path = "./templates/demo.docx";
//        handles(path);
    }
}
