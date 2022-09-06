package guave;

import com.google.common.collect.Table;


public class Application {
    public static void main(String[] args) {
        Excel3000 excel = new Excel3000();

        excel.setCell("A0", "5.0");
        excel.setCell("B0", "=$A0 * 2");
        excel.setCell("A1", "3.0");
        excel.setCell("B1", "=$A0 + $A1 +2");
        excel.setCell("A2", "=$A1^2 / 9");
        excel.setCell("B2", "=$A0 ^ $A1");
        System.out.println(excel.getTable().cellSet());
        Excel3000 excel2 = excel.evaluate();
        Table t = excel2.getTable();
        System.out.println(t.cellSet());
        excel2.writeExcel("test.xlsx");





    }
}
