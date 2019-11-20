package oneplay.office4j.utils;

import java.util.ArrayList;
import java.util.List;

/**
 * 对于二维区域，按照从上往下、从左往右的顺序遍历
 * 表达式1：(A1,B2)
 * 表达式2：(A1:C3)
 * 表达式3：(A1:C3,D1:F2)
 */
public class ReferenceUtil {

    public static List<String> getCellList(String reference) {
        List<String> cellList = new ArrayList<>();
        String[] regionList = reference.toUpperCase().split(",");
        for (String region : regionList) {
            if (region.contains(":")) {
                String[] tmp = region.split(":");
                cellList.addAll(getCelList(tmp[0], tmp[1]));
            } else {
                cellList.add(region);
            }
        }
        return cellList;
    }

    private static List<String> getCelList(String from, String to) {
        int fromCol = colFromStringToInt(from.replaceAll("[0-9]+", ""));
        int toCol = colFromStringToInt(to.replaceAll("[0-9]+", ""));
        int fromRow = Integer.valueOf(from.replaceAll("[A-Z]+", ""));
        int toRow = Integer.valueOf(to.replaceAll("[A-Z]+", ""));
        List<String> cellList = new ArrayList<>();
        for (int i = fromCol; i <= toCol; i++) {
            for (int j = fromRow; j <= toRow; j++) {
                cellList.add(colFromIntToString(i) + j);
            }
        }
        return cellList;
    }

    private static int colFromStringToInt(String col) {
        int index = 1;
        char[] colCharArray = col.toCharArray();
        for (int i = 0; i < colCharArray.length; i++) {
            index = 26 * i + colCharArray[i] - 'A' + 1;
        }
        return index;
    }

    private static String colFromIntToString(int index) {
        StringBuilder col = new StringBuilder();
        int remainder = index % 26;
        for (int quotient = index / 26; quotient > 0; quotient = quotient / 26) {
            col.append("Z");
        }
        if (remainder > 0) {
            col.append((char)('A' - 1 + remainder));
        }
        return col.toString();
    }

}
