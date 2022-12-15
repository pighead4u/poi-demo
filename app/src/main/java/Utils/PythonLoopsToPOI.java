package Utils;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class PythonLoopsToPOI {
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
//                            int ctpFirstValidPara = findFirstValidParagraph(cellToPlace);
//                            int ctpFirstValidRun = findFirstValidRun(cellToPlace.getParagraphs().get(ctpFirstValidPara));

//                            double cellToPlaceFontSize = cellToPlace.getParagraphs().get(ctpFirstValidPara).getRuns().
//                                    get(ctpFirstValidRun).getFontSizeAsDouble();
//                            String cellToPlaceFontFamily = cellToPlace.getParagraphs().get(ctpFirstValidPara).getRuns().
//                                    get(ctpFirstValidRun).getFontFamily();
                            XWPFRun cellToPlaceNewRun = cellToPlace.getParagraphs().get(cellToPlace.getParagraphs().size() - 1).createRun();
                            cellToPlaceNewRun.setText("{" + nationalStdMatcher.group(2) + "}");
                            cellToPlaceNewRun.setFontSize(FONT_SIZE);
                            cellToPlaceNewRun.setFontFamily(FONT_FAMILY);

                            XWPFTableCell currCell = table.getRow(i).getCell(0);
//                            int currCellFirstValidPara = findFirstValidParagraph(cellToPlace);
//                            int currCellFirstValidRun = findFirstValidRun(cellToPlace.getParagraphs().
//                                    get(currCellFirstValidPara));
//                            double currCellFontSize = currCell.getParagraphs().get(currCellFirstValidPara).getRuns().
//                                    get(currCellFirstValidRun).getFontSizeAsDouble();
//                            String currCellFontFamily = currCell.getParagraphs().get(currCellFirstValidPara).getRuns().
//                                    get(currCellFirstValidRun).getFontFamily();

                            for (int r = currCell.getParagraphs().size() - 1; r >= 0; r--) {
                                currCell.removeParagraph(r);
                            }
                            XWPFRun currCellNewRun = currCell.addParagraph().createRun();
                            currCellNewRun.setText("[" + nationalStdMatcher.group(1) + "]");
                            currCellNewRun.setFontSize(FONT_SIZE);
                            currCellNewRun.setFontFamily(FONT_FAMILY);

                        }
                        else if (loopMatcher.find()) { // deals with the case if the cell contains the Python-style for loop
//                            System.out.println("row info group 0: " + loopMatcher.group(0));
//                            System.out.println("row info group 1: " + loopMatcher.group(1));
//                            System.out.println("row info group 2: " + loopMatcher.group(2));
//                            System.out.println("row info group 3: " + loopMatcher.group(3));
//                            row.getTableCells().forEach(cell -> {
//                                System.out.println("cell contnt: ---------->>" + cell.getText());
//                            });

                            // REASONING: There are two sorts of content to place in the row above the for loop:
                            // one with a dot like "hwg_gtczjhdcl.instrument", and the other without a dot like
                            // "dts_kgg_jydzsy". We only want the tag so we can iterate the items in a POI-tl api
                            // fashion, so when "." is present, we'd take the tag before it.
                            String loopContentToPlace = !loopMatcher.group(2).contains(".") ? loopMatcher.group(2) :
                                    loopMatcher.group(3);


                            System.out.println("table text ----------" + table.getText());

                            System.out.println(table.getRows().size());

                            /* check if there exists previous row */
                            XWPFTableRow prevRow;

                            // TODO: FIX ADDING A NEW ROW ISSUE FOR demo.docx AT PAGE 41
                            if (i - 1 < 0) {
                                table.insertNewTableRow(0);
                                prevRow = table.getRow(0);
                                prevRow.createCell().addParagraph().createRun();
                            }
                            else {
                                prevRow = table.getRow(i - 1);
                            }


                            System.out.println(table.getRows().size());


                            XWPFTableCell cellToPlace = prevRow.getCell(0);
                            System.out.println("cell context ++++++++++" + cellToPlace.getText());

                            System.out.println(prevRow.getTableCells().get(findFirstValidCell(prevRow)));


                            System.out.println("-=-=-=-=-==-=-=-=-=-=-==-=-=-=-=-=-=-==-=-=-===-=-=-=-=");
//                            int ctpFirstValidPara = findFirstValidParagraph(cellToPlace);
//                            int ctpFirstValidRun = findFirstValidRun(cellToPlace.getParagraphs().get(ctpFirstValidPara));
//
//
//
//

//
//                            System.out.println("is valid para ------------- " + ctpFirstValidPara);
//                            System.out.println("is valid cell ============= " + ctpFirstValidRun);

//                            double fontSize = cellToPlace.getParagraphs().get(ctpFirstValidPara).getRuns().
//                                    get(ctpFirstValidRun).getFontSizeAsDouble();
//                            String fontFamily = cellToPlace.getParagraphs().get(ctpFirstValidPara).getRuns().
//                                    get(ctpFirstValidRun).getFontFamily();
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
//        ^(?=(\d)km(\d+)

        String s1 = "{{@dts_kgg_jdldcjsy_item.picpath}}";
        Pattern p1 = Pattern.compile(multiFieldsInOneCellRegex);
        Matcher m1 = p1.matcher(s1);

    }

    public static void main(String[] args) {


        String path = "./templates/demo.docx";
//        File file = new File(path);
//        File[] files = file.listFiles();
//        if (files != null) {
//            for (File f : files) {
//                System.out.println(f.getPath());
//                handles(f.getPath());
//            }
//        }

        handles(path);
    }
}
