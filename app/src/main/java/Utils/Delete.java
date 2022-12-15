package Utils;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class Delete {
    static final String TABLE_NAME = "2检测结果汇总";
    static final String FILE_PATH = "./templates/demo.docx";
    static final Set<String> EXCLUDE_STRINGS = new HashSet<>(Arrays.asList("t_lxzkcl", "t_gbwssy", "t_wssy"));

    /**
     * extract matched lines by regex
     *
     * @param tableName      the table to operate
     * @param targetFilePath the file path at which the docx file is
     * @param targetStrings  the strings that can be extracted to keep the lines
     */
    public static void extract(String tableName, String targetFilePath, Set<String> targetStrings) {
        try (XWPFDocument srcDoc = new XWPFDocument(new FileInputStream(targetFilePath))) {
            List<XWPFTable> tables = srcDoc.getTables();

            // get the table number
            int tableNum = 0;
            for (int i = 0; i < tables.size(); i++) {
                if (tables.get(i).getRow(0).getCell(0).getText().equals(tableName)) {
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
                    if (matcher.find() && targetStrings.contains(matcher.group(1))) {
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
            System.out.println("rows remain: " + table.getRows().size());

        } catch (IOException e) {
            System.out.println("Bad file path");
        }
    }
}


