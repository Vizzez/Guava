package guave;


import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Table;
import net.objecthunter.exp4j.Expression;
import net.objecthunter.exp4j.ExpressionBuilder;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Excel3000 {
    static final String REGEX = "^[A-Z]\\d+";
    //    static final String REGEX2 = "((\\$[A-Z]\\d+)?|(\\d)+)([\\+\\-\\*%^]?(($[A-Z]\\d+)?|(\\d)+))*";
    static final String REGEX2 = "((\\$[A-Z]\\d+)?|(\\d)+)(([\\+\\-\\*%^\\/]?)((\\$[A-Z]\\d+)?|(\\d)+))*";
    static final String REGEX3 = "\\$[A-Z]\\d+";
    private final Table<Integer, Integer, String> table;


    public Excel3000() {
        this.table = HashBasedTable.create();
    }

    private static int getNumRow(String coding) {
        return Integer.parseInt(coding.replaceAll("[^0-9]", ""));
    }

    private static int getLetterCol(String s) {
        return Character.getNumericValue(s.charAt(0)) - 10;
    }

    public void setCell(int r, int c, String v) {
        table.put(r, c, v.replaceAll("\\s+", ""));
    }

    public void setCell(String coding, String value) {

        if (coding.matches(REGEX)) {
            table.put(getNumRow(coding), getLetterCol(coding), value.replaceAll("\\s+", ""));
        } else System.out.println("Failed to set cell. Please use pattern :" + REGEX);
    }

    public String getCellAt(int row, int column) {
        return table.get(row, column);
    }

    public String getCellAt(String cell) {
        if (cell.matches(REGEX)) return table.get(getNumRow(cell), getLetterCol(cell));
        return null;
    }

    public Excel3000 evaluate() {
        Excel3000 excel = new Excel3000();
        Table tab = excel.getTable();
        Set<Table.Cell<Integer, Integer, String>> cellSet = table.cellSet();
        for (Table.Cell<Integer, Integer, String> cell : cellSet) {
            if (cell.getValue().charAt(0) == '=') {
                String[] split = cell.getValue().split("=");
                if (split[1] != null && split[1].matches(REGEX2)) {
                    Pattern pattern = Pattern.compile(REGEX3);
                    Matcher matcher = pattern.matcher(split[1]);
                    while (matcher.find()) {
                        String substr = matcher.group().substring(1);
                        split[1] = split[1].replace(matcher.group(), getCellAt(substr));
                    }
                    Expression exp = new ExpressionBuilder(split[1]).build();
                    tab.put(cell.getRowKey(), cell.getColumnKey(), Double.toString(exp.evaluate()));
                }
            } else {
                tab.put(cell.getRowKey(), cell.getColumnKey(), cell.getValue());
            }
        }
        return excel;
    }

    public Table<Integer, Integer, String> getTable() {
        return table;
    }

    public void writeExcel(String name) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        for (int i = 0; i < table.rowKeySet().size(); i++) {
            Row row = sheet.createRow(i);

            for (int j = 0; j < table.row(i).size(); j++) {
                Cell cell = row.createCell(j);
                System.out.println(table.get(i, j));
                cell.setCellValue(table.get(i, j));
                System.out.printf("row: %s col: %s val:%s%n", i, j, table.get(i, j));
            }

        }
        try (FileOutputStream outputStream = new FileOutputStream(name)) {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
