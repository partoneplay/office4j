package oneplay.office4j.test.word;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import oneplay.office4j.utils.Log4j2Util;
import oneplay.office4j.word.ChartOps;
import oneplay.office4j.word.Docx;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.junit.Test;

import java.io.File;

public class ChartOpsTests {
    private Logger logger = LogManager.getLogger(ChartOpsTests.class);

    @Test
    public void test() throws Exception {
        Log4j2Util.setLoggerLevel(ChartOps.class, "debug");

        logger.info("图表数据测试");
        String jsonData = FileUtils.readFileToString(new File("examples/Chart_Data.json"), "UTF-8");
        JSONArray jsonCharts = JSONObject.parseArray(jsonData);
        Docx docx = new Docx(new File("examples/Chart_Template.docx"));
        for (int i = 0; i < jsonCharts.size(); i++) {
            JSONObject jsonChart = jsonCharts.getJSONObject(i);
            String id = jsonChart.getString("id");
            Chart chart = docx.getChartByRId(id);
            logger.info("Chart Type is " + ChartOps.getType(chart));
            logger.info("The Original Data is ");
            ChartOps.print(chart);
            ChartOps.writeData(chart, jsonChart);
        }
        docx.save(new File("examples/Chart_Output.docx"));
    }

}
