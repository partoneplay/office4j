package oneplay.office4j.word;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.TypeReference;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.finders.TableFinder;
import org.docx4j.model.datastorage.migration.VariablePrepare;
import org.docx4j.model.fields.merge.DataFieldName;
import org.docx4j.model.fields.merge.MailMerger;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.DrawingML.Chart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.wml.Tbl;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Docx {
    private static final Logger logger = LogManager.getLogger(Docx.class);
    private WordprocessingMLPackage word;
    private RelationshipsPart relationshipsPart;

    private static final String TABLE = "table";
    private static final String CHART = "chart";
    private static final String MAPPING = "mapping";

    public Docx(File file) throws Exception {
        word = Docx4J.load(file);
        VariablePrepare.prepare(word);
        relationshipsPart = word.getMainDocumentPart().relationships;
    }

    public void writeData(String jsonStr) throws Docx4JException {
        JSONObject jsonDocx = JSONObject.parseObject(jsonStr);
        if (jsonDocx.containsKey(TABLE)) {
            logger.info("Writing Tables ...");
            JSONArray jsonTables = jsonDocx.getJSONArray(TABLE);
            List<Tbl> tblList = getTableList();
            int size = Math.min(jsonTables.size(), tblList.size());
            for (int i = 0; i < size; i++) {
                logger.info("Table index " + i + "/" + size);
                TableOps.writeData(tblList.get(i), jsonTables.getJSONArray(i));
            }
            logger.info("Writing Tables Done");
        }
        if (jsonDocx.containsKey(CHART)) {
            logger.info("Writing Charts ...");
            JSONArray jsonCharts = jsonDocx.getJSONArray(CHART);
            for (int i = 0; i < jsonCharts.size(); i++) {
                try {
                    JSONObject jsonChart = jsonCharts.getJSONObject(i);
                    String id = jsonChart.getString("id");
                    logger.info("Chart id " + id);
                    ChartOps.writeData(getChartByRId(id), jsonChart);
                } catch (Exception e) {
                    logger.error(e.getMessage(), e);
                }
            }
            logger.info("Writing Charts Done");
        }
        if (jsonDocx.containsKey(MAPPING)) {
            logger.info("Replacing Variables ...");
            Map<String, String> map = jsonDocx.getObject(MAPPING, new TypeReference<Map<String, String>>() {
                // nothing
            });
            logger.debug(map);
            textReplace(map);
            logger.info("Replacing Variables Done");
        }
        logger.info("Writing Data Done");
    }

    public void save(File file) throws Docx4JException {
        word.save(file);
    }

    private void textReplace(Map<String, String> mapping) throws Docx4JException {
        Map<DataFieldName, String> fieldMapping = new HashMap<>(mapping.size());
        for (Map.Entry<String, String> entry: mapping.entrySet()) {
            fieldMapping.put(new DataFieldName(entry.getKey()), entry.getValue());
        }
        MailMerger.setMERGEFIELDInOutput(MailMerger.OutputField.REMOVED);
        MailMerger.performMerge(word, fieldMapping, false);
    }

    public List<Chart> getChartList() throws Docx4JException {
        List<Chart> chartList = new ArrayList<>();
        for (Map.Entry<PartName, Part> entry : word.getParts().getParts().entrySet()) {
            String name = entry.getKey().getName();
            Part part = entry.getValue();
            if (part instanceof Chart) {
                logger.debug(name + " - " + ChartOps.getType((Chart) part));
                chartList.add((Chart) part);
            } else if (name.matches("chart[\\w]+\\.xml$")) {
                logger.warn("Unrecognized Chart " + name);
            }
        }
        return chartList;
    }

    public Chart getChartByRId(String rId) {
        return (Chart) relationshipsPart.getPart(rId);
    }

    public List<Tbl> getTableList() {
        TableFinder tableFinder = new TableFinder();
        new TraversalUtil(word.getMainDocumentPart(), tableFinder);
        return tableFinder.tblList;
    }

}
