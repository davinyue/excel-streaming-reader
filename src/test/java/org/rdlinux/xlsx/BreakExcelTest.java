package org.rdlinux.xlsx;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;

public class BreakExcelTest {
    @Test
    public void test() throws Exception {
        //InputStream is = Files.newInputStream(new File("D:/可疑数据/新建文件夹/不符情况.xlsx").toPath());
        InputStream is = Files.newInputStream(new File("src/test/resources/break.xlsx").toPath());
        Workbook workbook = StreamingReader.builder()
                .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
                .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                .open(is);            // InputStream or File for XLSX file (required)
        int sheetNum = 1;
        for (Sheet sheet : workbook) {
            System.out.println(sheet.getSheetName());
            int rowNUm = 1;
            for (Row r : sheet) {
                StringBuilder rowData = new StringBuilder();
                for (Cell c : r) {
                    rowData.append(c.getStringCellValue()).append(",");
                }
                System.out.println("第" + sheetNum + "个表, 第" + rowNUm + "行数据" + rowData);
                rowNUm++;
            }
            sheetNum++;
        }
    }
}
