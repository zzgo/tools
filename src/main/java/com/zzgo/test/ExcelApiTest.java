package com.zzgo.test;

import com.zzgo.util.excel.ExcelApi;

import java.util.*;

public class ExcelApiTest {

    //public void test() {
    //    String excelPath = "D:\\Desktop\\excelApi.xlsx";
    //    String[] titles = new String[]{"ID", "Name", "Sex", "Email", "Tel", "Address", "Note"};
    //    Map<Integer, Map<Integer, String>> values = new HashMap<Integer, Map<Integer, String>>();
    //    List<String> list = new ArrayList<String>();
    //    list.add("张三");
    //    list.add("李四");
    //    list.add("王五");
    //    list.add("李武");
    //    list.add("路飞");
    //    list.add("刘艳");
    //    list.add("荷花");
    //    list.add("才华");
    //    list.add("嘎子");
    //    list.add("你是谁");
    //    for (int i = 0; i < list.size(); i++) {
    //        Map<Integer, String> map = new HashMap<Integer, String>();
    //        String name = list.get(i);
    //        map.put(0, i + "");
    //        map.put(1, name);
    //        map.put(2, ran() > 2000 ? "男" : "女");
    //        map.put(3, ran() + "@qq.com");
    //        map.put(4, ran() + "");
    //        map.put(5, ran() + "（真实有效）");
    //        map.put(6, ran() + "（备注）");
    //        values.put(i, map);
    //    }
    //    ExcelApi ea = new ExcelApi(excelPath, titles, values);
    //    ea.write();
    //}
    public void test() {
        String excelPath = "D:\\Desktop\\excelApi.xlsx";
        String[] titles = new String[]{"ID", "Name", "Sex", "Email", "Tel", "Address", "Note"};
        Map<Integer, Map<Integer, String>> values = new HashMap<Integer, Map<Integer, String>>();
        for (int i = 0; i < 1000000; i++) {
            Map<Integer, String> map = new HashMap<Integer, String>();
            map.put(0, i + "");
            for (int j = 0; j < titles.length; j++) {
                map.put(j + 1, ran() + "");
            }
            values.put(i, map);
        }
        ExcelApi ea = new ExcelApi(excelPath, titles, values);
        ea.write();
    }

    public int ran() {
        return new Random().nextInt(100) + 1000;
    }

    public static void main(String[] args) {
        new ExcelApiTest().test();

    }
}
