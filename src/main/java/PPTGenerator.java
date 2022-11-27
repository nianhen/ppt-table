import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.sl.usermodel.TextParagraph;
import org.apache.poi.xslf.usermodel.*;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * @description: 生成表格
 * @author: nianhen
 * @since: 2022-11-27 10:23
 **/
public class PPTGenerator {

    private volatile static PPTGenerator generator = null;

    private PPTGenerator() {
    }

    public static PPTGenerator getInstance() {
        if (generator == null) {
            synchronized (PPTGenerator.class) {
                if (generator == null) {
                    generator = new PPTGenerator();
                }
            }
        }
        return generator;
    }


    public String generate(JSONArray tableData, String pptName) {
        XMLSlideShow ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();
        XSLFTextBox textBox = slide.createTextBox();
        textBox.setText("There are the TableData: ");
        textBox.setAnchor(new Rectangle(10, 10, 500, 20));

        XSLFTable table = slide.createTable();
        // 表头
        XSLFTableRow header = table.addRow();
        int groupCount = tableData.size();
        // 分组渲染
        for (int i = 0; i < groupCount; i++) {
            // 表头设值
            header.addCell().setText("Name");
            header.addCell().setText("Num");
            // 填入数据
            JSONArray lineGroup = tableData.getJSONArray(i);
            renderData(table, 1, lineGroup);
        }
        // 总计行
        XSLFTableRow totalRow = table.addRow();
        for (int i = 0; i < groupCount * 2; i++) {
            totalRow.addCell();
        }
        if (groupCount > 1) {
            totalRow.mergeCells(0, groupCount - 1);
            totalRow.mergeCells(groupCount, groupCount * 2 - 1);
        }
        List<XSLFTableCell> cells = totalRow.getCells();
        cells.get(0).setText("Total");
        cells.get(0).getTextParagraphs().get(0).setTextAlign(TextParagraph.TextAlign.CENTER);
        cells.get(groupCount).setText("1000"); // 实际业务场景传入或计算
        cells.get(groupCount).getTextParagraphs().get(0).setTextAlign(TextParagraph.TextAlign.CENTER);
        table.setAnchor(new Rectangle2D.Double(20, 90, 700, 500));
        return savePPT(ppt, pptName);
    }

    private int renderData(XSLFTable table, int rowNo, JSONArray lineGroup) {
        for (int j = 0; j < lineGroup.size(); j++) {
            List<XSLFTableRow> rows = table.getRows();
            JSONObject lineData = lineGroup.getJSONObject(j);
            if (rows.size() <= rowNo) {
                table.addRow();
            }
            XSLFTableRow lineRow = rows.get(rowNo++);
            lineRow.addCell().setText(lineData.getString("name"));
            lineRow.addCell().setText(lineData.getString("num"));
            // 产品域
            JSONArray areas = lineData.getJSONArray("children");
            if (areas != null && !areas.isEmpty()) {
                rowNo = renderData(table, rowNo, areas);
            }
        }
        return rowNo;
    }

    private String savePPT(XMLSlideShow ppt, String pptName) {
        String realName = pptName.endsWith(".pptx") ? pptName : pptName.concat(".pptx");
        try (FileOutputStream out = new FileOutputStream(realName)) {
            ppt.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        return System.getProperties().getProperty("user.dir").concat(File.pathSeparator).concat(realName);
    }
}
