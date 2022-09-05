package guave;

import com.google.common.collect.Table;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Application {
    public static void main (String[] args) {
        Excel3000 excel = new Excel3000();

        excel.setCell("A0","5");
        excel.setCell("B0","=$A0 * 2");
        excel.setCell("A1","3");
        excel.setCell("B1","=$A0 + $A1 +2");
        excel.setCell("A2","=$A1^2 / 9");
        System.out.println(excel.getTable().cellSet());
        Excel3000 excel2 = excel.evaluate();
        Table t=excel2.getTable();
        System.out.println(t.cellSet());


    }
}
