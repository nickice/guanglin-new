package com.guanglin.pptGen.datasource.excel;

import com.guanglin.pptGen.datasource.DataSourceBase;
import com.guanglin.pptGen.exception.DataSourceException;
import com.guanglin.pptGen.exception.ExcelValidationException;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.Project;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by pengyao on 03/06/2017.
 */
public abstract class ExcelDataSourceBase extends DataSourceBase {

    protected final static int titleRowIndex = 2;

    protected final static Map<String, String> fields = new HashMap<String, String>();

    static {
        fields.put("A3", "编号");
        fields.put("B3", "项目名称");
        fields.put("C3", "区域");
        fields.put("D3", "位置描述");
        fields.put("E3", "社区分类");
        fields.put("F3", "租售均价");
        fields.put("G3", "社区居住规模");
        fields.put("H3", "入住率");
        fields.put("I3", "各社区内受众描述");
        fields.put("J3", "楼层");
        fields.put("K3", "门洞数");
        fields.put("L3", "电梯总数");
        fields.put("M3", "等候厅数");
        fields.put("N3", "合同数");
        fields.put("O3", "实际数");
        fields.put("P3", "楼号细分");
        fields.put("Q3", "广告发布日期");
        fields.put("R3", "监测日期");
    }

    protected Workbook workbook;


    public ExcelDataSourceBase(Project project, FileInputStream excelInputStream)
            throws IOException, InvalidFormatException {
        super(project);
        this.workbook = WorkbookFactory.create(excelInputStream);
    }

    @Override
    public Project mapProjectItemData() throws DataSourceException {

        try {
            this.project.setItems(extractItemsFromWorkbook());
            return this.project;
        } catch (ExcelValidationException ex) {
            throw new DataSourceException(ex);
        }
    }

    protected abstract List<Item> extractItemsFromWorkbook() throws ExcelValidationException, DataSourceException;

}
