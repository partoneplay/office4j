package oneplay.office4j.word;

import com.alibaba.fastjson.JSONArray;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.docx4j.XmlUtils;
import org.docx4j.wml.*;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

public class TableOps {
    private static final Logger logger = LogManager.getLogger(TableOps.class);

    /**
     * 逐行填充数据
     *
     * @param tbl       Tbl
     * @param jsonTable JSONArray
     */
    public static void writeData(Tbl tbl, JSONArray jsonTable) {
        int nRow = jsonTable.size();
        List<Tr> trList = getTrList(tbl, nRow);
        for (int i = 0; i < jsonTable.size(); i++) {
            JSONArray jsonRow = jsonTable.getJSONArray(i);
            List<Tc> tcList = getTcList(trList.get(i), jsonRow.size());
            for (int j = 0; j < jsonRow.size(); j++) {
                writeTc(tcList.get(j), jsonRow.getString(j));
            }
        }
    }

    private static List<Tr> getTrList(Tbl tbl, int nRow) {
        List<Object> content = tbl.getContent();
        List<Object> trObjectList = content.stream().filter(x -> XmlUtils.unwrap(x) instanceof Tr).collect(Collectors.toList());
        autoExpand(content, trObjectList, nRow);
        return tbl.getContent().stream().map(XmlUtils::unwrap).filter(x -> x instanceof Tr).map(x -> (Tr) x).collect(Collectors.toList());
    }

    private static List<Tc> getTcList(Tr tr, int nCol) {
        List<Object> content = tr.getContent();
        List<Object> tcObjectList = content.stream().filter(x -> XmlUtils.unwrap(x) instanceof Tc).collect(Collectors.toList());
        autoExpand(content, tcObjectList, nCol);
        return tr.getContent().stream().map(XmlUtils::unwrap).filter(x -> x instanceof Tc).map(x -> (Tc) x).collect(Collectors.toList());
    }

    private static void writeTc(Tc tc, String value) {
        logger.debug("Cell Value: " + value);
        // 一个单元格可以有多行数据
        List<String> pValueList = Arrays.asList(value.split("\n", -1));
        int pValueListSize = pValueList.size();
        List<Object> content = tc.getContent();
        List<Object> pObjectList = tc.getContent().stream().map(XmlUtils::unwrap).filter(x -> x instanceof P).collect(Collectors.toList());
        autoExpand(content, pObjectList, pValueListSize);

        List<P> pList = tc.getContent().stream().map(XmlUtils::unwrap).filter(x -> x instanceof P).map(x -> (P) x).collect(Collectors.toList());
        // 无法保留单元格内行内格式信息
        for (int i = 0; i < pValueListSize; i++) {
            String pValue = pValueList.get(i);
            logger.debug("Cell line Value: " + pValue);
            P p = pList.get(i);
            p.getContent().clear();
            Text text = new Text();
            text.setValue(pValue);
            R r = new R();
            r.getContent().add(text);
            p.getContent().add(r);
        }
    }

    /**
     * 自动扩展行或列，不足则后补，多余则删除
     */
    private static void autoExpand(List<Object> content, List<Object> objectList, int n) {
        assert !objectList.isEmpty() && n > 0;

        int objectListSize = objectList.size();
        if (objectListSize < n) {
            // 不足，复制最后一个追加
            Object object = objectList.get(objectListSize - 1);
            for (int i = objectListSize; i < n; i++) {
                content.add(XmlUtils.deepCopy(object));
            }
        } else if (objectListSize > n) {
            // 多余自动删除
            for (int i = objectListSize - 1; i >= n; i--) {
                content.remove(objectList.get(i));
            }
        }
    }

}
