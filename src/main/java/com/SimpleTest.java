package com;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.IOException;

public class SimpleTest {
    public static void main(String[] args) throws IOException {
        // 建立工作簿
        Workbook workbook = new XSSFWorkbook(new FileInputStream("src/main/java/com/sample.xlsx"));
        // 取得第一個工作表
        Sheet sheet = workbook.getSheetAt(0);
        // 取得第一行
        Row row = sheet.getRow(0);
        // 取得第一個欄位
        Cell cell = row.getCell(0);
        // 輸出資料
        System.out.println(cell.getStringCellValue());
        workbook.close();
    }
}