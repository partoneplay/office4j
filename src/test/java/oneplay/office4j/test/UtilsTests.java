package oneplay.office4j.test;

import oneplay.office4j.utils.ReferenceUtil;
import org.junit.Test;

public class UtilsTests {

    @Test
    public void referenceTest() {
        String[] references = new String[]{"A1,B2,C3", "A1:C3", "A1:B2,E4:G8"};
        for (String reference : references) {
            for (String cell : ReferenceUtil.getCellList(reference)) {
                System.out.println(cell);
            }
            System.out.println();
        }
    }

}
