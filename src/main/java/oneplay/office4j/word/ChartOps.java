package oneplay.office4j.word;

import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.TypeReference;
import org.docx4j.dml.chart.*;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.slf4j.LoggerFactory;

import javax.xml.bind.JAXBException;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Office 图表操作类
 * 散点图、气泡图需要特殊处理，其他类型图表可通用处理
 *
 * @author oneplay
 */
public class ChartOps {
    private static org.slf4j.Logger logger = LoggerFactory.getLogger(ChartSer.class);

    private static final String SERIES = "series";
    private static final String MAPPING = "mapping";

    /**
     * 逐个序列填充数据
     *
     * @param chart     Chart
     * @param jsonChart JSONObject
     * @throws Docx4JException Docx4JException
     */
    public static void writeData(Chart chart, JSONObject jsonChart) throws Docx4JException, JAXBException {
        String chartName = chart.getPartName().getName();
        logger.info("Writing " + chartName + " ...");
        if (jsonChart.containsKey(MAPPING)) {
            // 除格式化外一版不会出现Split Run问题
            chart.variableReplace(jsonChart.getObject(MAPPING, new TypeReference<Map<String, String>>() {
                // nothing
            }));
        }
        if (jsonChart.containsKey(SERIES)) {
            ChartSer chartSer = getChartSer(chart);
            chartSer.writeData(jsonChart.getJSONObject(SERIES));
        }
        logger.info("Writing Chart " + chartName + " Done.");
    }

    /**
     * 获取指定图表绘图区中的序列元素
     * CTAreaChart 面积图
     * CTArea3DChart 三维面积图
     * CTBarChart 条形图
     * CTBar3DChart 三维条形图
     * CTBubbleChart 气泡图
     * CTDoughnutChart 圆环图
     * CTLineChart 折线图
     * CTLine3DChart 三位折线图
     * CTOfPieChart 复合饼图
     * CTPieChart 饼图
     * CTPie3DChart 三维饼图
     * CTRadarChart 雷达图
     * CTScatterChart 散点图
     * CTStockChart 股价图
     * CTSurfaceChart 曲面图
     * CTSurface3DChart 三维曲面图
     *
     * @param chart 图表对象Chart
     * @return ChartSer
     */
    private static ChartSer getChartSer(Chart chart) throws Docx4JException {
        ChartSer chartSer = new ChartSer(chart);
        for (Object ctChart : chart.getContents().getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart()) {
            String ctChartClass = ctChart.getClass().getSimpleName();
            switch (ctChartClass) {
                case "CTAreaChart":
                    List<CTAreaSer> CTAreaChartList = ((CTAreaChart) ctChart).getSer();
                    chartSer.addCat(CTAreaChartList.stream().map(CTAreaSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTAreaChartList.stream().map(CTAreaSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTAreaChartList.stream().map(CTAreaSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTArea3DChart":
                    List<CTAreaSer> CTArea3DChartList = ((CTArea3DChart) ctChart).getSer();
                    chartSer.addCat(CTArea3DChartList.stream().map(CTAreaSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTArea3DChartList.stream().map(CTAreaSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTArea3DChartList.stream().map(CTAreaSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTBarChart":
                    List<CTBarSer> CTBarChartList = ((CTBarChart) ctChart).getSer();
                    chartSer.addCat(CTBarChartList.stream().map(CTBarSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTBarChartList.stream().map(CTBarSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTBarChartList.stream().map(CTBarSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTBar3DChart":
                    List<CTBarSer> CTBar3DChartList = ((CTBar3DChart) ctChart).getSer();
                    chartSer.addCat(CTBar3DChartList.stream().map(CTBarSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTBar3DChartList.stream().map(CTBarSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTBar3DChartList.stream().map(CTBarSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTBubbleChart":
                    List<CTBubbleSer> CTBubbleChartList = ((CTBubbleChart) ctChart).getSer();
                    chartSer.addCat(CTBubbleChartList.stream().map(CTBubbleSer::getXVal).collect(Collectors.toList()));
                    chartSer.addTx(CTBubbleChartList.stream().map(CTBubbleSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTBubbleChartList.stream().map(CTBubbleSer::getYVal).collect(Collectors.toList()));
                    chartSer.addSize(((CTBubbleChart) ctChart).getSer().stream().map(CTBubbleSer::getBubbleSize).collect(Collectors.toList()));
                    break;
                case "CTDoughnutChart":
                    List<CTPieSer> CTDoughnutChartList = ((CTDoughnutChart) ctChart).getSer();
                    chartSer.addCat(CTDoughnutChartList.stream().map(CTPieSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTDoughnutChartList.stream().map(CTPieSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTDoughnutChartList.stream().map(CTPieSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTLineChart":
                    List<CTLineSer> CTLineChartList = ((CTLineChart) ctChart).getSer();
                    chartSer.addCat(CTLineChartList.stream().map(CTLineSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTLineChartList.stream().map(CTLineSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTLineChartList.stream().map(CTLineSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTLine3DChart":
                    List<CTLineSer> CTLine3DChartList = ((CTLine3DChart) ctChart).getSer();
                    chartSer.addCat(CTLine3DChartList.stream().map(CTLineSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTLine3DChartList.stream().map(CTLineSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTLine3DChartList.stream().map(CTLineSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTOfPieChart":
                    List<CTPieSer> CTOfPieChartList = ((CTOfPieChart) ctChart).getSer();
                    chartSer.addCat(CTOfPieChartList.stream().map(CTPieSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTOfPieChartList.stream().map(CTPieSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTOfPieChartList.stream().map(CTPieSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTPieChart":
                    List<CTPieSer> CTPieChartList = ((CTPieChart) ctChart).getSer();
                    chartSer.addCat(CTPieChartList.stream().map(CTPieSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTPieChartList.stream().map(CTPieSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTPieChartList.stream().map(CTPieSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTPie3DChart":
                    List<CTPieSer> CTPie3DChartList = ((CTPie3DChart) ctChart).getSer();
                    chartSer.addCat(CTPie3DChartList.stream().map(CTPieSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTPie3DChartList.stream().map(CTPieSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTPie3DChartList.stream().map(CTPieSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTRadarChart":
                    List<CTRadarSer> CTRadarChartList = ((CTRadarChart) ctChart).getSer();
                    chartSer.addCat(CTRadarChartList.stream().map(CTRadarSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTRadarChartList.stream().map(CTRadarSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTRadarChartList.stream().map(CTRadarSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTScatterChart":
                    List<CTScatterSer> CTScatterChartList = ((CTScatterChart) ctChart).getSer();
                    chartSer.addCat(CTScatterChartList.stream().map(CTScatterSer::getXVal).collect(Collectors.toList()));
                    chartSer.addTx(CTScatterChartList.stream().map(CTScatterSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTScatterChartList.stream().map(CTScatterSer::getYVal).collect(Collectors.toList()));
                    break;
                case "CTStockChart":
                    List<CTLineSer> CTStockChartList = ((CTStockChart) ctChart).getSer();
                    chartSer.addCat(CTStockChartList.stream().map(CTLineSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTStockChartList.stream().map(CTLineSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTStockChartList.stream().map(CTLineSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTSurfaceChart":
                    List<CTSurfaceSer> CTSurfaceChartList = ((CTSurfaceChart) ctChart).getSer();
                    chartSer.addCat(CTSurfaceChartList.stream().map(CTSurfaceSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTSurfaceChartList.stream().map(CTSurfaceSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTSurfaceChartList.stream().map(CTSurfaceSer::getVal).collect(Collectors.toList()));
                    break;
                case "CTSurface3DChart":
                    List<CTSurfaceSer> CTSurface3DChartList = ((CTSurface3DChart) ctChart).getSer();
                    chartSer.addCat(CTSurface3DChartList.stream().map(CTSurfaceSer::getCat).collect(Collectors.toList()));
                    chartSer.addTx(CTSurface3DChartList.stream().map(CTSurfaceSer::getTx).collect(Collectors.toList()));
                    chartSer.addVal(CTSurface3DChartList.stream().map(CTSurfaceSer::getVal).collect(Collectors.toList()));
                    break;
                default:
                    logger.warn("Unsupported Chart Type " + ctChartClass);
            }
        }
        return chartSer;
    }

    public static String getType(Chart chart) throws Docx4JException {
        List<Object> chartList = chart.getContents().getChart().getPlotArea().getAreaChartOrArea3DChartOrLineChart();
        if (chartList.isEmpty()) {
            logger.error(chart.getPartName().getName() + " has no Chart");
            return "";
        }
        return chartList.stream().map(x -> x.getClass().getSimpleName()).collect(Collectors.joining(","));
    }

    public static void print(Chart chart) throws Docx4JException {
        ChartSer chartSer = getChartSer(chart);
        chartSer.printData();
    }

}
