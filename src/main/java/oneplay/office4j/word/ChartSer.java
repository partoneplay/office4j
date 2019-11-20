package oneplay.office4j.word;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import oneplay.office4j.utils.ReferenceUtil;
import org.docx4j.dml.chart.*;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.io3.Save;
import org.docx4j.openpackaging.packages.SpreadsheetMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.SpreadsheetML.WorksheetPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EmbeddedPackagePart;
import org.docx4j.utils.BufferUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xlsx4j.sml.CTRst;
import org.xlsx4j.sml.Cell;
import org.xlsx4j.sml.STCellType;

import java.io.ByteArrayOutputStream;
import java.util.*;
import java.util.stream.Collectors;

class ChartSer {
    private static Logger logger = LoggerFactory.getLogger(ChartSer.class);
    private static final String CATEGORY = "category";
    private static final String DATA = "data";
    private static final String NAME = "name";
    private static final String VALUE = "value";
    private static final String X = "x";
    private static final String SIZE = "size";

    private static final String REFERENCE_REGEX = "(Sheet[0-9]+)|[!$)(]";

    private EmbeddedPackagePart excelPart;
    private SpreadsheetMLPackage excel;
    private Map<String, Cell> cellMap;
    private List<CTRst> sharedStringList;

    ChartSer(Chart chart) throws Docx4JException {
        String excelPartName = chart.relationships.getRelationshipByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package").getTarget();
        excelPart = (EmbeddedPackagePart) chart.getPackage().getParts().get(new PartName(excelPartName.replace("..", "/word")));
        this.excel = SpreadsheetMLPackage.load(BufferUtil.newInputStream(excelPart.getBuffer()));
        // 与图表对应的嵌入的excel文件只有一个sheet页
        String workSheetPartName = excel.getWorkbookPart().relationships.getRelationshipByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet").getTarget();
        WorksheetPart worksheetPart = (WorksheetPart) excel.getParts().get(new PartName("/xl/" + workSheetPartName));
        // table.xml 用在excel内定义表格，和普通区域不同；对于图表多余，需要删除，否则会造成部分数据不一致，word打开提示文件损坏
        worksheetPart.relationships.removeRelationshipsByType("http://schemas.openxmlformats.org/officeDocument/2006/relationships/table");
        worksheetPart.getContents().getTableParts().getTablePart().clear();
        // sharedStrings.xml 当字符串类数据超过1个字符，真实数据会存放在此xml文件中，sheet通过索引引用
        this.sharedStringList = excel.getWorkbookPart().getSharedStrings().getContents().getSi();
        List<Cell> cellList = worksheetPart.getContents().getSheetData().getRow().stream().flatMap(row -> Arrays.stream(row.getC().toArray(new Cell[0]))).collect(Collectors.toList());
        cellMap = new HashMap<>(cellList.size());
        for (Cell cell : cellList) {
            String r = cell.getR();
            if (!cellMap.containsKey(r)) {
                cellMap.put(r, cell);
            }
        }
    }

    /**
     * 序列类别
     * 对于散点图序列和气泡图序列，对应xVAl
     * 对于其他类型图表序列，对应cat
     * 有strRef/strCache或者numRef/numCache两种情况
     */
    private List<CTAxDataSource> catList = new ArrayList<>();

    /**
     * 序列名称
     * 对应tx
     * 只有strRef/strCache 一种情况
     */
    private List<CTSerTx> txList = new ArrayList<>();

    /**
     * 序列值
     * 对于散点图序列和气泡图序列，对应yVAl
     * 对于其他类型图表序列，对应val
     * 只有numRef/numCache 一种情况
     */
    private List<CTNumDataSource> valList = new ArrayList<>();

    /**
     * 气泡大小
     * 对于气泡图序列，对应bubbleSize
     * 对于其他类型图表序列无效
     * 只有numRef/numCache 一种情况
     */
    private List<CTNumDataSource> sizeList = new ArrayList<>();

    void addCat(List<CTAxDataSource> list) {
        catList.addAll(list);
    }

    void addTx(List<CTSerTx> list) {
        txList.addAll(list);
    }

    void addVal(List<CTNumDataSource> list) {
        valList.addAll(list);
    }

    void addSize(List<CTNumDataSource> list) {
        sizeList.addAll(list);
    }

    /**
     * 写入替换更新数据
     */
    void writeData(JSONObject jsonChart) throws Docx4JException {
        writeCat(jsonChart);
        writeTx(jsonChart);
        writeVal(jsonChart);
        writeSize(jsonChart);
        // 保存嵌入与图表对应的excel的数据
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        Save excelSave = new Save(excel);
        excelSave.save(byteArrayOutputStream);
        excelPart.setBinaryData(byteArrayOutputStream.toByteArray());
    }

    private void writeCat(JSONObject jsonChart) {
        if (catList.isEmpty()) {
            return;
        }
        logger.debug("Writing Cat ...");
        List<List<Object>> ctStrValLists = catList.stream().map(ChartSer::mergeCatVal).collect(Collectors.toList());
        if (ctStrValLists.isEmpty()) {
            return;
        }
        List<String> referenceList = catList.stream().map(ChartSer::mergeCatF).collect(Collectors.toList());
        if (jsonChart.containsKey(CATEGORY)) {
            // 不支持次要横坐标轴的图表
            List<String> categoryList = jsonChart.getJSONArray(CATEGORY).stream().map(Object::toString).collect(Collectors.toList());
            for (List<Object> ctStrValList : ctStrValLists) {
                writeOneSeries(ctStrValList, referenceList.get(0), categoryList);
            }
        } else {
            boolean isX = false;
            JSONArray jsonSeries = jsonChart.getJSONArray(DATA);
            for (int i = 0; i < jsonSeries.size(); i++) {
                JSONObject jsonObject = jsonSeries.getJSONObject(i);
                if (jsonObject.containsKey(X)) {
                    isX = true;
                    break;
                }
            }
            if (isX) {
                // 散点图或者气泡图
                writeSeries(ctStrValLists, referenceList, jsonSeries, X);
            }
        }
    }

    private void writeTx(JSONObject jsonChart) {
        if (txList.isEmpty()) {
            return;
        }
        logger.debug("Writing Tx ...");
        List<List<CTStrVal>> ctStrValLists = txList.stream().map(ChartSer::mergeTxVal).collect(Collectors.toList());
        if (ctStrValLists.isEmpty()) {
            return;
        }
        List<String> referenceList = txList.stream().map(ChartSer::mergeTxF).collect(Collectors.toList());
        JSONArray jsonSeries = jsonChart.getJSONArray(DATA);
        writeSeries(ctStrValLists, referenceList, jsonSeries, NAME);
    }

    private void writeVal(JSONObject jsonChart) {
        if (valList.isEmpty()) {
            return;
        }
        logger.debug("Writing Val ...");
        List<List<CTNumVal>> ctStrValLists = valList.stream().map(ChartSer::mergeNumVal).collect(Collectors.toList());
        if (ctStrValLists.isEmpty()) {
            return;
        }
        List<String> referenceList = valList.stream().map(ChartSer::mergeValF).collect(Collectors.toList());
        JSONArray jsonSeries = jsonChart.getJSONArray(DATA);
        writeSeries(ctStrValLists, referenceList, jsonSeries, VALUE);
    }

    private void writeSize(JSONObject jsonChart) {
        if (sizeList.isEmpty()) {
            return;
        }
        logger.debug("Writing Size ...");
        List<List<CTNumVal>> ctStrValLists = sizeList.stream().map(ChartSer::mergeNumVal).collect(Collectors.toList());
        if (ctStrValLists.isEmpty()) {
            return;
        }
        List<String> referenceList = sizeList.stream().map(ChartSer::mergeValF).collect(Collectors.toList());
        JSONArray jsonSeries = jsonChart.getJSONArray(DATA);
        writeSeries(ctStrValLists, referenceList, jsonSeries, SIZE);
    }

    private <T> void writeSeries(List<List<T>> ctValLists, List<String> referenceList, JSONArray jsonSeries, String key) {
        if (ctValLists.size() != jsonSeries.size()) {
            logger.warn("Series Num not the same as the Data Num.");
        }
        int size = Math.min(ctValLists.size(), jsonSeries.size());
        for (int i = 0; i < size; i++) {
            JSONObject jsonObject = jsonSeries.getJSONObject(i);
            if (jsonObject.containsKey(key)) {
                List<T> ctStrValList = ctValLists.get(i);
                List<String> dataList = jsonObject.getJSONArray(key).stream().map(Object::toString).collect(Collectors.toList());
                writeOneSeries(ctStrValList, referenceList.get(i), dataList);
            }
        }
    }

    private <T> void writeOneSeries(List<T> ctValList, String reference, List<String> data) {
        int ctValSize = ctValList.size();
        int dataSize = data.size();
        int size = Math.min(ctValSize, dataSize);
        if (ctValSize != dataSize) {
            logger.warn("CTStrVal Series Size not the same as the Data Size."
                    + "The Series is : " + ctValList.stream().map(ChartSer::getV).collect(Collectors.toList()).toString() + ". "
                    + "The Data is : " + data.toString() + ".");
        }
        List<String> cellNameList = ReferenceUtil.getCellList(reference);
        for (int i = 0; i < size; i++) {
            String v = data.get(i);
            setV(ctValList.get(i), v);
            // 更新嵌入与图表对应的excel的数据
            String cellName = cellNameList.get(i);
            if (cellMap.containsKey(cellName)) {
                Cell cell = cellMap.get(cellName);
                if (cell.getT() == STCellType.S) {
                    // 数据在sharedStrings.xml中
                    int index = Integer.valueOf(cell.getV());
                    sharedStringList.get(index).getT().setValue(v);
                } else {
                    cell.setV(v);
                }
                logger.debug("Set " + cellName + " to " + v);
            }
        }
    }

    private static String getV(Object ctVal) {
        if (ctVal instanceof CTStrVal) {
            return ((CTStrVal) ctVal).getV();
        } else if (ctVal instanceof CTNumVal) {
            return ((CTNumVal) ctVal).getV();
        } else {
            logger.error("cannot apply getV for class of type " + ctVal.getClass().getSimpleName());
            return "";
        }
    }

    private static void setV(Object ctVal, String v) {
        if (ctVal instanceof CTStrVal) {
            ((CTStrVal) ctVal).setV(v);
        } else if (ctVal instanceof CTNumVal) {
            ((CTNumVal) ctVal).setV(v);
        } else {
            logger.error("cannot apply setV for class of type " + ctVal.getClass().getSimpleName());
        }
    }

    private static List<Object> mergeCatVal(CTAxDataSource ctAxDataSource) {
        List<Object> tmpList = new ArrayList<>();
        if (ctAxDataSource.getStrRef() != null) {
            tmpList.addAll(ctAxDataSource.getStrRef().getStrCache().getPt());
        } else if (ctAxDataSource.getNumRef() != null) {
            tmpList.addAll(ctAxDataSource.getNumRef().getNumCache().getPt());
        } else {
            logger.error("There is no ref and cache from CTAxDataSource. The Cat Data should be referred, not specified.");
        }
        return tmpList;
    }

    private static List<CTStrVal> mergeTxVal(CTSerTx ctSerTx) {
        List<CTStrVal> tmpList = new ArrayList<>();
        if (ctSerTx.getStrRef() != null) {
            tmpList.addAll(ctSerTx.getStrRef().getStrCache().getPt());
        } else {
            logger.error("There is no ref and cache from CTSerTx. The Tx Data should be referred, not specified.");
        }
        return tmpList;
    }

    private static List<CTNumVal> mergeNumVal(CTNumDataSource ctNumDataSource) {
        List<CTNumVal> tmpList = new ArrayList<>();
        if (ctNumDataSource.getNumRef() != null) {
            tmpList.addAll(ctNumDataSource.getNumRef().getNumCache().getPt());
        } else {
            logger.error("There is no ref and cache from CTNumDataSource. The Val Data should be referred, not specified.");
        }
        return tmpList;
    }

    private static String mergeCatF(CTAxDataSource ctAxDataSource) {
        String f = "";
        if (ctAxDataSource.getStrRef() != null) {
            f = ctAxDataSource.getStrRef().getF();
        } else if (ctAxDataSource.getNumRef() != null) {
            f = ctAxDataSource.getNumRef().getF();
        } else {
            logger.error("There is no ref and cache from CTAxDataSource. The Cat Data should be referred, not specified.");
        }
        return f.replaceAll(REFERENCE_REGEX, "");
    }

    private static String mergeTxF(CTSerTx ctAxDataSource) {
        String f = "";
        if (ctAxDataSource.getStrRef() != null) {
            f = ctAxDataSource.getStrRef().getF();
        } else {
            logger.error("There is no ref and cache from CTAxDataSource. The Cat Data should be referred, not specified.");
        }
        return f.replaceAll(REFERENCE_REGEX, "");
    }

    private static String mergeValF(CTNumDataSource ctAxDataSource) {
        String f = "";
        if (ctAxDataSource.getNumRef() != null) {
            f = ctAxDataSource.getNumRef().getF();
        } else {
            logger.error("There is no ref and cache from CTAxDataSource. The Cat Data should be referred, not specified.");
        }
        return f.replaceAll(REFERENCE_REGEX, "");
    }

    void printData() {
        List<List<Object>> catLists = catList.stream().map(ChartSer::mergeCatVal).collect(Collectors.toList());
        List<List<CTStrVal>> txLists = txList.stream().map(ChartSer::mergeTxVal).collect(Collectors.toList());
        List<List<CTNumVal>> valLists = valList.stream().map(ChartSer::mergeNumVal).collect(Collectors.toList());
        List<List<CTNumVal>> sizeLists = sizeList.stream().map(ChartSer::mergeNumVal).collect(Collectors.toList());
        int size = catLists.size();
        for (int i = 0; i < size; i++) {
            System.out.println(txLists.get(i).stream().map(ChartSer::getV).collect(Collectors.toList()));
            System.out.println(catLists.get(i).stream().map(ChartSer::getV).collect(Collectors.toList()));
            System.out.println(valLists.get(i).stream().map(ChartSer::getV).collect(Collectors.toList()));
            if (!sizeLists.isEmpty()) {
                System.out.println(sizeLists.get(i).stream().map(ChartSer::getV).collect(Collectors.toList()));
            }
        }
    }

}
