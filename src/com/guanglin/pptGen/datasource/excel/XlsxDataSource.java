package com.guanglin.pptGen.datasource.excel;

import com.google.common.base.Strings;
import com.guanglin.pptGen.exception.DataSourceException;
import com.guanglin.pptGen.exception.ExcelValidationException;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.Project;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * Created by pengyao on 03/06/2017.
 */
public class XlsxDataSource extends ExcelDataSourceBase {
    private XSSFWorkbook xssfWorkbook;

    public XlsxDataSource(Project project, FileInputStream fileInputStream)
            throws IOException, InvalidFormatException {
        super(project, fileInputStream);

        // instance the xssfworkbook
        this.xssfWorkbook = (XSSFWorkbook) super.workbook;
    }

    @Override
    protected List<Item> extractItemsFromWorkbook() throws ExcelValidationException, DataSourceException {

        // check the workbook contains the required fields and values
        ValidateWorkbookContent();

        Sheet sheet = xssfWorkbook.getSheetAt(0);
        List<Item> itemsRet = new ArrayList<>();

        // the first data row equals title row index plus 1
        int validRowIndex = titleRowIndex + 1;

        do {
            Row row = sheet.getRow(validRowIndex);

            // Convert the String to int successfully, then start to map the value to item
            if (row.getCell(0).getNumericCellValue() > 0) {

                Map<String, String> tmpMap = new HashMap<>();
                Item item = new Item();
                for (Map.Entry<String, String> e : fields.entrySet()) {

                    // foreach cells to get the values
                    for (int i = 0; i < fields.size(); i++) {
                        Cell cell = row.getCell(i);
                        switch (cell.getCellTypeEnum()) {
                            case STRING:
                                tmpMap.put(e.getValue(), cell.getStringCellValue());
                                break;
                            case BOOLEAN:
                                tmpMap.put(e.getValue(), Boolean.valueOf(cell.getBooleanCellValue()).toString());
                                break;
                            case NUMERIC:
                                tmpMap.put(e.getValue(), Double.valueOf(cell.getNumericCellValue()).toString());
                                break;
                            case _NONE:
                            case BLANK:
                                tmpMap.put(e.getValue(), "");
                                break;
                            default:
                                throw new DataSourceException("unknown data format.");
                        }
                    }
                    // set the map
                    item.setFields(tmpMap);
                }
                itemsRet.add(item);
                validRowIndex++;
            } else {
                validRowIndex = 0;
            }

        } while (validRowIndex > 0);

        return itemsRet;
    }

    private void ValidateWorkbookContent() throws ExcelValidationException {

        // 1. validate there should be only one sheet
        if (xssfWorkbook.getNumberOfSheets() != 1)
            throw new ExcelValidationException("这个Excel文件存在多个Sheet", "应该每个文件只包含一个Sheet");

        // 2. validate the first sheet isn't the default name
        if (Strings.isNullOrEmpty(xssfWorkbook.getSheetAt(0).getSheetName())
                && xssfWorkbook.getSheetAt(0).getSheetName().equals("Sheet1")) {
            throw new ExcelValidationException("Sheet的名字不正确", "sheet的名字应该不是Sheet1");
        }

        Sheet sheet = xssfWorkbook.getSheetAt(0);
        // 3. validate field names, the 3th row should be field name row
        Row row = sheet.getRow(titleRowIndex);
        for (int i = 1; i <= fields.size(); i++) {
            Cell cell = row.getCell(i - 1);
            if (!Strings.isNullOrEmpty(cell.getStringCellValue())) {

                // Ascii code of 'A' is 65
                String key = String.format("%s3", (char) (64 + i));

                if (fields.get(key) == null
                        || !cell.getStringCellValue().contains(fields.get(key))) {
                    throw new ExcelValidationException(
                            String.format("字段[%s]名称[%s]不正确", key, cell.getStringCellValue()),
                            String.format("字段[%s]应该是[%s]", key, fields.get(key))
                    );

                }
            }

        }
        // 4. validate the 4th row first cell should be coverted to Integer
        Cell a4Cell = sheet.getRow(3).getCell(0);
        try {
            if (a4Cell == null || a4Cell.getNumericCellValue() < 1) {
                throw new ExcelValidationException("没有一个有效的广告数据", "在第4行应该至少有一条有效的广告数据");
            }
        } catch (NumberFormatException ex) {
            throw new ExcelValidationException("没有一个有效的广告数据", "在第4行应该至少有一条有效的广告数据");
        }
    }
}
