package com.idaoben.commom.excel;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author Sherman
 * created in 2018/8/24
 */

public class NoHeaderExcel extends AbstractExcel {
    private final static boolean hasHeader = false;
    private HSSFSheet sheet;

    public NoHeaderExcel(HSSFSheet sheet) {
        super(sheet);
        this.sheet = sheet;
    }

    public NoHeaderExcel(String sheetName, String excelPath) {
        HSSFWorkbook wb = null;
        try {
            wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(excelPath)));
        } catch (IOException e) {
            e.printStackTrace();
        }
        sheet = sheetName == null || sheetName.isEmpty() ? wb.getSheetAt(0) : wb.getSheet(sheetName);
    }


    @Override
    public List<Map<String, String>> getPayload() {
        List<Map<String, String>> payLoad = new ArrayList<>();
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            HSSFRow currentRow = sheet.getRow(i);
            Map<String, String> map = new HashMap<>();
            for (int j = 0; j <= sheet.getRow(i).getLastCellNum(); j++) {
                map.put(String.valueOf(j), getStringFormatCellValue(currentRow.getCell(j)));
            }
            payLoad.add(map);
        }
        return payLoad;
    }


}
