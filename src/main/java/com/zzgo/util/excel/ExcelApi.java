package com.zzgo.util.excel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Map;

/**
 * <p>excel的读写操作工具</p>
 */
public class ExcelApi {
    private String excelPath;
    private String[] titles;
    private Map<Integer, Map<Integer, String>> values;
    private Workbook wb;

    public ExcelApi(String excelPath, String[] titles, Map<Integer, Map<Integer, String>> values) {
        this.excelPath = excelPath;
        this.values = values;
        this.titles = titles;
        init();
    }


    private void init() {
        try {
            if (excelPath.endsWith("xls")) {
                wb = new HSSFWorkbook();
            } else if (excelPath.endsWith("xlsx")) {
                wb = new SXSSFWorkbook();
            } else {
                throw new Exception("文件格式不正确！");
            }
        } catch (Exception e) {

        }
    }

    public void write() {
        long startTime = System.currentTimeMillis();
        try {
            Sheet sheet = (Sheet) wb.createSheet("sheet1");
            //表头
            Row titleRow = (Row) sheet.createRow(0);
            Cell cell = titleRow.createCell(0);
            for (int i = 0; i < titles.length; i++) {
                cell = titleRow.createCell(i);
                cell.setCellValue(titles[i]);
                sheet.setColumnWidth(i, 20 * 256);
            }
            titleRow.setHeight((short) 540);
            for (int i = 0; i < values.size(); i++) {
                Row row = (Row) sheet.createRow(i + 1);
                Map<Integer, String> colValues = values.get(i);
                for (int j = 0; j < colValues.size(); j++) {
                    row.createCell(j).setCellValue(colValues.get(j));
                }
            }
            //创建文件流
            OutputStream stream = new FileOutputStream(excelPath);
            //写入数据
            wb.write(stream);
            //关闭文件流
            stream.close();
            System.out.println("用时=" + (System.currentTimeMillis() - startTime));
        } catch (Exception e) {

        }
    }
}
