package oneplay.office4j.test.word;

import oneplay.office4j.utils.Log4j2Util;
import oneplay.office4j.word.ChartOps;
import oneplay.office4j.word.Docx;
import org.apache.commons.io.FileUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.junit.Test;

import java.io.File;

public class DocxTests {
    private Logger logger = LogManager.getLogger(DocxTests.class);

    @Test
    public void test() throws Exception {
        Log4j2Util.setLoggerLevel(Docx.class, "debug");

        logger.info("文档数据测试");
        String jsonData = FileUtils.readFileToString(new File("examples/Docx_Data.json"), "UTF-8");
        Docx docx = new Docx(new File("examples/Docx_Template.docx"));
        docx.writeData(jsonData);
        docx.save(new File("examples/Docx_Output.docx"));
    }

    @Test
    public void printAllChartType() throws Exception {
        Docx docx = new Docx(new File("examples/Chart.docx"));
        for (Chart chart: docx.getChartList()) {
            logger.info(chart.getPartName().getName() + " " + ChartOps.getType(chart));
        }
    }

}
