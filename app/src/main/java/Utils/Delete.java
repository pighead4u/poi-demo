package Utils;


import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class Delete {
    static final String FILE_PATH = "./templates/demo.docx";
    static final Set<String> EXCLUDE_STRINGS = new HashSet<String>(Arrays.asList("t_lxzkcl", "t_gbwssy", "t_wssy"));

    public static void extract(String targetFilePath, Set<String> targetStrings) throws Exception {
        XWPFDocument srcDoc = new XWPFDocument(new FileInputStream(FILE_PATH));

        List<XWPFParagraph> paragraphs = srcDoc.getParagraphs();
        List<XWPFTable> tables = srcDoc.getTables();
        XWPFTable table = tables.get(5); // the table 2检测结果汇总
        List<XWPFTableRow> rows = table.getRows();

        // create patterns for regex matching
        String p1 = "t_lxzkcl|t_gbwssy|t_wssy";

        // find the rows that do not contain the pattern
        Pattern pattern = Pattern.compile(p1);
        List<Integer> rowIndices = new ArrayList<>();

        // create a list to store all possible outcomes
        List<Integer> indices = IntStream.range(2, rows.size()).boxed().collect(Collectors.toList());

        // save the wanted lines from the list
        for (int i = 2; i < rows.size(); i++) { // since the contents starts at index 2
            for (XWPFTableCell cell : rows.get(i).getTableCells()) {
                Matcher matcher = pattern.matcher(cell.getText());
                if (matcher.find()) {
                    indices.remove((Integer) i);
                    break;
                }
            }
        }

        Collections.reverse(indices);

        // keep only the rows that match the pattern
        for (int i : indices) {
            table.removeRow(i);
        }
        System.out.println("rows remain: " + table.getRows().size());


//        for (XWPFTableRow row : rows) {
//            for (XWPFTableCell cell : row.getTableCells()) {
//                System.out.println(cell.getText());
//            }
//            System.out.println("\n");
//        }
    }

    public static void main(String[] args) throws Exception {
        extract(FILE_PATH, EXCLUDE_STRINGS);
    }
}


