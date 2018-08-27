package com.idaoben.commom.excel;

import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * Unit test for simple App.
 */
public class AppTest {
    /**
     * 测试建表，写表操作
     */
    @Test
    public void testExportExcel() {
        //测试数据
        String[] headers = new String[]{"A", "B", "C", "D", "E"};
        List<List<Object>> valueList = new LinkedList<>();
        for (char i = 'A'; i <= 'E'; i++) {
            List<Object> rowList = new LinkedList<>();
            for (int j = 0; j <= 4; j++) {
                rowList.add(i + String.valueOf(j));
            }
            valueList.add(rowList);
        }

        AbstractExcel excel = new ExcelBuilder("报名表")
                .header(headers)
                .content(valueList)
                .autoColumnWidth()
                .build(true);

        try {
            File file = new File("D:\\ideawork\\excelcommom\\src\\main\\resourses\\test.xls");
            FileOutputStream op = new FileOutputStream(file);
            excel.write(op);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 测试读取表数据操作
     */
    @Test
    public void testImportExcel() {
        AbstractExcel excel = new HeaderExcel(null, "D:\\ideawork\\excelcommom\\src\\main\\resourses\\test.xls");
        List<Map<String, String>> values = excel.getPayload();
        values.forEach(stringStringMap -> {
            stringStringMap.entrySet().forEach(stringStringEntry -> {
                System.out.println(stringStringEntry.getKey() + "---->" + stringStringEntry.getValue());
            });

        });
    }

}
