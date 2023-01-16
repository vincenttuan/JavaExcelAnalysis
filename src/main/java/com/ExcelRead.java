package com;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class ExcelRead {
    public static List<ExcelBean> read(String filePath) throws IOException {
        List<ExcelBean> list = new ArrayList<>();
        
        // 建立工作簿
        Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
        // 取得第一個工作表
        Sheet sheet = workbook.getSheetAt(0);
        
        for(int r=4;r<=23;r++) {
            Row row = sheet.getRow(r);
            ExcelBean bean =  new ExcelBean();
            for(int i=2;i<=4;i++) {
                Cell cell = row.getCell(i);
                // 輸出資料
                switch (i) {
                    case 2: //字串格式
                        String branch = cell.getStringCellValue();
                        bean.setBranch(branch);
                        break;
                    case 3: // 日期格式
                        Date date = cell.getDateCellValue();
                        bean.setDate(date);
                        break;
                    case 4: //字串格式
                        String reason = cell.getStringCellValue().toUpperCase().replaceAll(" ", "");
                        bean.setReason(reason);
                        break;    
                }
            }
            list.add(bean);
        }
        workbook.close();
        return list;
    }
    
    public static Set<String> getDeviceIds(List<ExcelBean> list) {
        Set<String> deviceIds = new LinkedHashSet<>();
        list.forEach(e -> {
            String reason = e.getReason();
            if(reason != null) {
                String[] array = reason.split("機型/設備序號:");
                if(array.length > 1) {
                    String deviceId = array[1];
                    deviceId = deviceId.substring(0, deviceId.indexOf("\n"));
                    if(deviceId.contains("#")) {
                        deviceId = deviceId.substring(0, deviceId.indexOf("#"));
                    }
                    if(deviceId.contains("*")) {
                        deviceId = deviceId.substring(0, deviceId.indexOf("*"));
                    }
                    deviceId = deviceId.trim();
                    deviceIds.add(deviceId);
                }
            }
        });
        return deviceIds;
    }
    
    public static List<Map<String, Object>> analysis(int year, String filePath) throws Exception{
        List<Map<String, Object>> report  = new ArrayList<>();
        List<ExcelBean> list = read(filePath);
        Set<String> deviceIds = getDeviceIds(list);
        Calendar calendar = Calendar.getInstance();
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        for (int m = 1; m <= 12; m++) {
            //System.out.println(m + " 月");
            for(ExcelBean bean : list) {
                calendar.setTime(bean.getDate());
                int yyyy = calendar.get(Calendar.YEAR);
                int mm = calendar.get(Calendar.MONTH) + 1;
                if(yyyy == year && mm == m) {
                    for(String deviceId : deviceIds) {
                        String reason = bean.getReason();
                        if(reason.contains(deviceId)) {
                            Map<String, Object> map = new LinkedHashMap<>();
                            map.put("date", bean.getDate());
                            map.put("branch", bean.getBranch().trim());
                            map.put("deviceId", deviceId.trim());
                            
                            //System.out.print(sdf.format(bean.getDate()) + ":" + bean.getBranch() + ":" + deviceId);
                            String star = deviceId + "*";
                            if(reason.contains(star)) {
                                int begin = reason.indexOf(star)+star.length()-1;
                                int end = begin + 2;
                                String count = reason.substring(begin, end);
                                //System.out.print(count);
                                count = count.replaceAll("\\*", "");
                                map.put("count", count);
                            } else {
                                map.put("count", "1");
                            }
                            //System.out.println();
                            report.add(map);
                        }
                    }
                }
            }
        }
        return report;
    }
}