package Utils;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLOutput;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class PythonLoopsToPOI {


    /**
     * Find the indices of tables that contain Python for loops
     * @return a list of indices of all tables that contain Python for loops
     */
    private static ArrayList<Integer> findTables() {

        return null;
    }

    /**
     * Find the indices of rows in a table that contain Python for loops
     * @return a list of indices of all rows that contain Python for loops
     */
    private static ArrayList<Integer> findRows() {

        return null;
    }

//    public static ArrayList<String> getCellStringsInNextLine(XWPFDocument doc, int tableNum, int rowNum) {
//        List<XWPFTable>
//    }

    public static void handles(String filePath) {
        try (XWPFDocument tempDoc = new XWPFDocument(new FileInputStream(filePath))) {


            List<XWPFTable> tables = tempDoc.getTables();
            // {%tr for item in dts_kgg_jydzsy %}
            String loopRegex = "%.*for.?(.*)in\\s((.*)\\.|.*).*%"; // {%tr for instrument in hwg_gtczjhdcl.instrument %}
            String endloopRegex = "%.*(endfor).*%"; // {%tr endfor %}
            String multiFieldsInOneCellRegex = "(@?)(?:.*\\.)(.*)}}";
            Pattern loopPattern = Pattern.compile(loopRegex);
            Pattern endloopPattern = Pattern.compile(endloopRegex);
            Pattern multiFieldsInOneCellPattern = Pattern.compile(multiFieldsInOneCellRegex);

            int tabacc = 0;
            for (XWPFTable table : tables) {
                System.out.println(tabacc);
                tabacc += 1;
                List<Integer> rowsToDelete = new ArrayList<>();
                for (int i = 0; i < table.getRows().size(); i++) {
                    XWPFTableRow row = table.getRow(i);
                    for (int j = 0; j < row.getTableCells().size(); j++) {
                        String cellText = row.getCell(j).getText();
                        Matcher loopMatcher = loopPattern.matcher(cellText); // match the Python-style for loop
                        Matcher endloopMatcher = endloopPattern.matcher(cellText);

                        if (loopMatcher.find()) { // deals with the case if the cell contains the Python-style for loop
                            System.out.println("row info group 0: " + loopMatcher.group(0));
                            System.out.println("row info group 1: " + loopMatcher.group(1));
                            System.out.println("row info group 2: " + loopMatcher.group(2));
                            System.out.println("row info group 3: " + loopMatcher.group(3));

                            // REASONING: There are two sorts of content to place in the row above the for loop:
                            // one with a dot like "hwg_gtczjhdcl.instrument", and the other without a dot like
                            // "dts_kgg_jydzsy". We only want the tag so we can iterate the items in a POI-tl api
                            // fashion, so when "." is present, we'd take the tag before it.
                            String loopContentToPlace = !loopMatcher.group(2).contains(".") ? loopMatcher.group(2) :
                                    loopMatcher.group(3);

                            XWPFTableRow prevRow = table.getRow(i - 1);
                            XWPFTableCell cellToPlace = prevRow.getCell(0);

                            double fontSize = cellToPlace.getParagraphs().get(0).getRuns().get(0).getFontSizeAsDouble();
                            String fontFamily = cellToPlace.getParagraphs().get(0).getRuns().get(0).getFontFamily();
                            XWPFRun newRun = cellToPlace.addParagraph().createRun();
                            newRun.setText(loopContentToPlace);
                            newRun.setFontFamily(fontFamily);
                            newRun.setFontSize(fontSize);


                            rowsToDelete.add(i); // saves the rows that contain "for ... in ..."
                            XWPFTableRow nextRow = table.getRow(i + 1);

                            // get each cell from the next row
                            for (XWPFTableCell cell : nextRow.getTableCells()) {
                                String nextRowCellText = cell.getText();
                                System.out.println(nextRowCellText);
                                Matcher multiFieldsInOneCellMatcher = multiFieldsInOneCellPattern.
                                        matcher(nextRowCellText);

                                if (multiFieldsInOneCellMatcher.find()) {
                                    System.out.println("group: ->>>>>>" + multiFieldsInOneCellMatcher.group());
                                    // TODO: complete the case where "@" (syb. for pictures) is present




                                    String cellMatcherResult = multiFieldsInOneCellMatcher.group();
                                    XWPFParagraph paragraph = cell.getParagraphs().get(0);
                                    XWPFRun run = paragraph.getRuns().get(0);

                                    // remove all old cells
                                    ParagraphAlignment paragraphAlignment = paragraph.getAlignment();

                                    cell.removeParagraph(0);

                                    // add new content to cells
                                    XWPFParagraph newParagraph = cell.addParagraph();
                                    newParagraph.createRun().setText("[" + cellMatcherResult + "]");
                                    newParagraph.getRuns().get(0).setFontFamily(fontFamily);
                                    newParagraph.getRuns().get(0).setFontSize(fontSize);
                                    newParagraph.setAlignment(paragraphAlignment);
                                }
                            }
                        }

                        if (endloopMatcher.find()) {
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
            FileOutputStream fos = new FileOutputStream("output.docx");
            tempDoc.write(fos);
        } catch (IOException e) {
            System.out.println("bad file");
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

        System.out.println(m1.find());
        System.out.println(m1.group());

    }

    public static void main(String[] args) {

        String path = "./templates/0.4kV电缆分支箱.docx";
//        foo();
        handles(path);
    }
}
