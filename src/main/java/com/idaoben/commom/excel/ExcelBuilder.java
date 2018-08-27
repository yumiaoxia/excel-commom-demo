package com.idaoben.commom.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.Date;
import java.util.List;

/**
 * 建表工具类
 *
 * @author Sherman
 * email:1253950375@qq.com
 * created in 2018/8/24
 */
@Slf4j
public class ExcelBuilder {

    private static HSSFSheet sheet;
    private static HSSFWorkbook wb;
    private static boolean hasHeader;

    /**
     * 初始化
     *
     * @param excellName 表名
     */
    public ExcelBuilder(String excellName) {
        wb = new HSSFWorkbook();
        sheet = wb.createSheet(excellName);
    }

    /**
     * 设置表头，装配表头数据
     *
     * @param value 字符串数组，用来作为表头的值
     */
    public ExcelBuilder header(String... value) {
        if (value != null && value.length != 0) {
            //设置表头样式
            HSSFCellStyle cellStyle = wb.createCellStyle();
            cellStyle.setFont(font("黑体", true, 12));
            HSSFRow row = sheet.createRow(0);
            for (int i = 0; i < value.length; i++) {
                HSSFCell cell = row.createCell(i);
                cell.setCellValue(value[i]);
                cell.setCellStyle(cellStyle);
            }
            hasHeader = true;
        }
        return this;
    }

    /**
     * excel 表内容装配
     *
     * @param content 待装配表格内容的二维数组
     * @return
     */
    public ExcelBuilder content(List<List<Object>> content) {
        if (content != null && !content.isEmpty()) {
            int index;
            for (int i = 0; i < content.size(); i++) {
                index = hasHeader == false ? i : i + 1;
                HSSFRow row = sheet.createRow(index);
                for (int j = 0; j < content.get(i).size(); j++) {
                    String r = "";
                    Object value = content.get(i).get(j);
                    //根据数据类型装配
                    if (value instanceof String) {
                        r = (String) value;
                    } else if (value instanceof Number) {
                        r = String.valueOf(value);
                    } else if (value instanceof BigDecimal) {
                        r = String.valueOf(value);
                    } else {
                        if (!(value instanceof Date) && !(value instanceof Timestamp)) {
                            if (!(value instanceof ZonedDateTime) && !(value instanceof LocalDateTime)) {
                                if (value instanceof Enum) {
                                    r = ((Enum) value).name();
                                } else if (value != null) {

                                    log.info("Error of create row, Unknow field type: " + value.getClass().getName());
                                }
                            } else {
                                DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM-dd");
                                r = formatter.format((TemporalAccessor) value);
                            }
                        } else {
                            DateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                            r = sdf.format(value);
                        }
                    }

                    row.createCell(j).setCellValue(r);
                }
            }
        }
        return this;
    }

    /**
     * 自动调整列宽大小
     */
    public ExcelBuilder autoColumnWidth() {
        for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
            int maxLength = 0;
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                String value = sheet.getRow(i).getCell(j).getStringCellValue();
                int length = 0;
                if (value != null) {
                    length = value.getBytes().length;
                }
                if (length > maxLength) {
                    maxLength = length;
                }
            }
            sheet.setColumnWidth(j, maxLength > 30 ? (30 * 256 + 186) : (maxLength * 256 + 186));
        }
        return this;
    }

    /**
     * 实例化
     *
     * @param hasHeader 是否有表头
     * @return Excel表格
     */
    public AbstractExcel build(Boolean hasHeader) {
        return hasHeader ? new HeaderExcel(sheet) : new NoHeaderExcel(sheet);
    }

    /**
     * @param fontName 字体名字
     * @param isBold   是否粗体
     * @param fontSize 字体大小
     * @return 字体
     */
    private HSSFFont font(String fontName, boolean isBold, int fontSize) {
        HSSFFont font = wb.createFont();
        if (fontName != null) font.setFontName(fontName);
        else font.setFontName("黑体");
        font.setBold(isBold);
        font.setFontHeightInPoints((short) fontSize);
        return font;
    }

}
