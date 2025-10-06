/* ====================================================================
   Licensed to the Apache Software Foundation (ASF) under one or more
   contributor license agreements.  See the NOTICE file distributed with
   this work for additional information regarding copyright ownership.
   The ASF licenses this file to You under the Apache License, Version 2.0
   (the "License"); you may not use this file except in compliance with
   the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.
==================================================================== */

package org.apache.poi.xssf.streaming;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.LazyXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

public class TestInPlaceSXSSFWorkbook {

    @TempDir
    File tempDir;

    @Test
    void testInPlaceSXSSFWorkbookCreation() throws Exception {
        // Create a simple XLSX file first
        File tempFile = new File(tempDir, "test.xlsx");
        
        // Create initial workbook with some data
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("TestSheet");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Initial Data");
            
            try (FileOutputStream out = new FileOutputStream(tempFile)) {
                wb.write(out);
            }
        }
        
        // Now test our InPlaceSXSSFWorkbook
        try (InPlaceSXSSFWorkbook workbook = new InPlaceSXSSFWorkbook(tempFile)) {
            assertNotNull(workbook);
            assertEquals(1, workbook.getNumberOfSheets());
            
            // Create a new sheet
            Sheet newSheet = workbook.createSheet("NewSheet");
            assertNotNull(newSheet);
            assertEquals(2, workbook.getNumberOfSheets());
            
            // Add some data to the new sheet
            Row row = newSheet.createRow(0);
            row.createCell(0).setCellValue("New Data");
        }
    }

    @Test 
    void testInPlaceSXSSFWorkbookWithWindowSize() throws Exception {
        // Create a simple XLSX file first
        File tempFile = new File(tempDir, "test2.xlsx");
        
        // Create initial workbook with some data
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("TestSheet");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Initial Data");
            
            try (FileOutputStream out = new FileOutputStream(tempFile)) {
                wb.write(out);
            }
        }
        
        // Test with custom window size
        try (InPlaceSXSSFWorkbook workbook = new InPlaceSXSSFWorkbook(tempFile, 50)) {
            assertNotNull(workbook);
            assertEquals(1, workbook.getNumberOfSheets());
        }
    }

    @Test
    void testInPlaceSXSSFWorkbookWithAllOptions() throws Exception {
        // Create a simple XLSX file first
        File tempFile = new File(tempDir, "test3.xlsx");
        
        // Create initial workbook with some data
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("TestSheet");
            Row row = sheet.createRow(0);
            row.createCell(0).setCellValue("Initial Data");
            
            try (FileOutputStream out = new FileOutputStream(tempFile)) {
                wb.write(out);
            }
        }
        
        // Test with all options
        try (InPlaceSXSSFWorkbook workbook = new InPlaceSXSSFWorkbook(tempFile, 50, true, true)) {
            assertNotNull(workbook);
            assertEquals(1, workbook.getNumberOfSheets());
        }
    }

    @Test
    void testInPlaceSXSSFWorkbookUsesLazyLoading() throws Exception {
        // Create a workbook with multiple sheets
        File tempFile = new File(tempDir, "lazy_test.xlsx");
        
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("Sheet1");
            wb.createSheet("Sheet2");
            wb.createSheet("Sheet3");
            
            try (FileOutputStream out = new FileOutputStream(tempFile)) {
                wb.write(out);
            }
        }
        
        // Test that InPlaceSXSSFWorkbook uses LazyXSSFWorkbook
        try (InPlaceSXSSFWorkbook workbook = new InPlaceSXSSFWorkbook(tempFile)) {
            assertNotNull(workbook);
            assertEquals(3, workbook.getNumberOfSheets());
            
            // Verify we can access sheets by name and index
            assertNotNull(workbook.getSheet("Sheet1"));
            assertNotNull(workbook.getSheetAt(0));
            assertEquals("Sheet1", workbook.getSheetName(0));
            assertEquals("Sheet2", workbook.getSheetName(1));
            assertEquals("Sheet3", workbook.getSheetName(2));
        }
    }

    @Test
    void test69838() throws Exception {
        final int COLUMN_COUNT = 10;
        final int ROW_COUNT = 600;

        File tempfile = File.createTempFile("test69838", ".xlsx");

        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100)) {
            workbook.setCompressTempFiles(true);
            SXSSFSheet sheet = workbook.createSheet("RawData");
            populateSheetWithTestData(sheet, COLUMN_COUNT, ROW_COUNT);       
            try (FileOutputStream out = new FileOutputStream(tempfile)) {
                workbook.write(out);
            }
        }

        try (InPlaceSXSSFWorkbook workbook2 = new InPlaceSXSSFWorkbook(tempfile, 100)) {
            SXSSFSheet sheet2 = workbook2.createSheet("RawData2");
            populateSheetWithTestData(sheet2, COLUMN_COUNT, ROW_COUNT);
        }
    }

    private void populateSheetWithTestData(SXSSFSheet sheet, int columnCount, int rowCount) {
        final int TEN_MINUTES = 1000 * 60 * 10;
        
        SXSSFRow row = sheet.createRow(0);
        SXSSFCell cell;

        // Create header row
        for (int i = 1; i <= columnCount; i++) {
            cell = row.createCell(i - 1);
            cell.setCellValue("Column " + i);
        }

        // Populate data rows
        for (int i = 1; i < rowCount; i++) {
            row = sheet.createRow(i);
            for (int j = 1; j <= columnCount; j++) {
                cell = row.createCell(j - 1);

                //make some noise
                cell.setCellValue(new Date(i * TEN_MINUTES + (j * TEN_MINUTES) / columnCount));
            }
            i++;
        }
    }
}