package com;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;

public class Main {
    public static void main(String[] args) throws Exception {
        String filePath = "src/main/java/com/data.xlsx";
        int year = 2022;
        List<Map<String, Object>> report = ExcelRead.analysis(year, filePath);
        
        Calendar calendar = Calendar.getInstance();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        for (int m = 1; m <= 12; m++) {
            System.out.println(m + " 月");
            for(Map<String, Object> map : report) {
                calendar.setTime((Date)map.get("date"));
                int yyyy = calendar.get(Calendar.YEAR);
                int mm = calendar.get(Calendar.MONTH) + 1;
                if(yyyy == year && mm == m) {
                    System.out.printf("日期:%10s 分公司:%-20s 設備ID:%-7s 台數:%2s\n",
                            sdf.format(map.get("date")), 
                            map.get("branch"), 
                            map.get("deviceId"), 
                            map.get("count"));
                }
            }
         }
    }   
}
