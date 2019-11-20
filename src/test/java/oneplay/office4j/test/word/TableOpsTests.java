package oneplay.office4j.test.word;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import oneplay.office4j.utils.Log4j2Util;
import oneplay.office4j.word.Docx;
import oneplay.office4j.word.TableOps;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.docx4j.wml.Tbl;
import org.junit.Test;

import java.io.File;
import java.util.List;

public class TableOpsTests {
    private static Logger logger = LogManager.getLogger(TableOpsTests.class);

    @Test
    public void test() throws Exception {
        Log4j2Util.setLoggerLevel(TableOps.class, "debug");

        logger.info("表格数据测试");
        String jsonData = FileUtils.readFileToString(new File("examples/Table_Data.json"), "UTF-8");
        JSONArray jsonTables = JSONObject.parseArray(jsonData);
        Docx docx = new Docx(new File("examples/Table_Template.docx"));
        List<Tbl> tblList = docx.getTableList();
        for (int i = 0; i < tblList.size(); i++) {
            TableOps.writeData(tblList.get(i), jsonTables.getJSONArray(i));
        }
        docx.save(new File("examples/Table_Output.docx"));
    }

}
