package oneplay.office4j.test;

import org.junit.Test;

import java.util.Arrays;
import java.util.List;

public class JSONTests {
    private String json1 = "{\"title\": \"图表名称\",\"category\":[\"类别1\",\"类别2\",\"类别3\",\"类别4\"],\"series\":[{\"name\":\"系列1\",\"value\":{\"类别1\":\"4.3\",\"类别2\":\"2.5\",\"类别3\":\"3.5\",\"类别4\":\"4.5\"}},{\"name\":\"系列2\",\"value\":{\"类别1\":\"2.4\",\"类别2\":\"4.4\",\"类别3\":\"1.8\",\"类别4\":\"2.8\"}},{\"name\":\"系列3\",\"value\":{\"类别1\":\"2\",\"类别2\":\"2\",\"类别3\":\"3\",\"类别4\":\"5\"}}]}";
    private String json2 = "{\"title\": \"图表名称\",\"series\":[{\"name\":\"X值\",\"value\":[\"0.7\",\"1.8\",\"2.6\"]},{\"name\":\"Y值\",\"value\":[\"2.7\",\"3.2\",\"0.8\"]},{\"name\":\"大小\",\"value\":[\"10\",\"4\",\"8\"]}]}";

    @Test
    public void test() {
        List<String> a = Arrays.asList("1", "2");
        System.out.println(a.toString());
    }

}