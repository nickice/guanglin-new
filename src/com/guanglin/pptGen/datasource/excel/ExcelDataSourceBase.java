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

    protected final static int titleRowIndex = 3;

    protected final static Map<String, String> fields = new HashMap<String, String>();

    static {
        fields.put("A", "编号");
        fields.put("B", "项目名称");
        fields.put("C", "区域");
        fields.put("D", "位置描述");
        fields.put("F", "社区分类");
        fields.put("E", "租售均价");
        fields.put("G", "社区居住规模");
        fields.put("H", "入住率");
        fields.put("I", "各社区受众描述");
        fields.put("J", "楼层");
        fields.put("K", "门洞数");
        fields.put("L", "电梯总数");
        fields.put("M", "等候厅数");
        fields.put("N", "合同数");
        fields.put("O", "实际数");
        fields.put("P", "楼号细分");
        fields.put("Q", "广告发布日期");
        fields.put("R", "检测日期");
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

    protected abstract List<Item> extractItemsFromWorkbook() throws ExcelValidationException;

}
