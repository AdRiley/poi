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

package org.apache.poi.xssf.usermodel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

public class TestLazyXSSFWorkbook {

    @TempDir
    File tempDir;

    @Test
    void testLazyLoadingBasicFunctionality() throws Exception {
        // Create a multi-sheet workbook
        File tempFile = new File(tempDir, "lazy_test.xlsx");
        
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            // Create multiple sheets with data
            Sheet sheet1 = wb.createSheet("Sheet1");
            Row row1 = sheet1.createRow(0);
            row1.createCell(0).setCellValue("Sheet 1 Data");
            
            Sheet sheet2 = wb.createSheet("Sheet2");
            Row row2 = sheet2.createRow(0);
            row2.createCell(0).setCellValue("Sheet 2 Data");
            
            Sheet sheet3 = wb.createSheet("Sheet3");
            Row row3 = sheet3.createRow(0);
            row3.createCell(0).setCellValue("Sheet 3 Data");
            
            try (FileOutputStream out = new FileOutputStream(tempFile)) {
                wb.write(out);
            }
        }
        
        // Test that LazyXSSFWorkbook works like regular workbook for now
        try (LazyXSSFWorkbook lazyWb = new LazyXSSFWorkbook(tempFile)) {
            // Basic functionality should work
            assertEquals(3, lazyWb.getNumberOfSheets());
            assertEquals(3, lazyWb.getNumberOfLoadedSheets()); // For current simple implementation
            
            // Test sheet name access
            assertEquals("Sheet1", lazyWb.getSheetName(0));
            assertEquals("Sheet2", lazyWb.getSheetName(1));
            assertEquals("Sheet3", lazyWb.getSheetName(2));
            
            // All sheets should be considered loaded in current implementation
            assertTrue(lazyWb.isSheetLoaded(0));
            assertTrue(lazyWb.isSheetLoaded(1));
            assertTrue(lazyWb.isSheetLoaded(2));
            
            // Access first sheet
            Sheet sheet1 = lazyWb.getSheetAt(0);
            assertNotNull(sheet1);
            assertEquals("Sheet1", sheet1.getSheetName());
            assertEquals("Sheet 1 Data", sheet1.getRow(0).getCell(0).getStringCellValue());
            
            // Access sheet by name
            Sheet sheet2 = lazyWb.getSheet("Sheet2");
            assertNotNull(sheet2);
            assertEquals("Sheet2", sheet2.getSheetName());
            assertEquals("Sheet 2 Data", sheet2.getRow(0).getCell(0).getStringCellValue());
        }
    }

    @Test
    void testLazyIterator() throws Exception {
        // Create a multi-sheet workbook
        File tempFile = new File(tempDir, "lazy_iterator_test.xlsx");
        
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            wb.createSheet("First");
            wb.createSheet("Second");
            wb.createSheet("Third");
            
            try (FileOutputStream out = new FileOutputStream(tempFile)) {
                wb.write(out);
            }
        }
        
        // Test iteration works correctly
        try (LazyXSSFWorkbook lazyWb = new LazyXSSFWorkbook(tempFile)) {
            assertEquals(3, lazyWb.getNumberOfLoadedSheets()); // All loaded in current implementation
            
            int count = 0;
            for (Sheet sheet : lazyWb) {
                assertNotNull(sheet);
                count++;
            }
            
            assertEquals(3, count);
            assertEquals(3, lazyWb.getNumberOfLoadedSheets());
        }
    }

    @Test
    void testComparisonWithRegularWorkbook() throws Exception {
        // Create a workbook with some data
        File tempFile = new File(tempDir, "comparison_test.xlsx");
        
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("TestSheet");
            for (int i = 0; i < 100; i++) {
                Row row = sheet.createRow(i);
                row.createCell(0).setCellValue("Row " + i);
                row.createCell(1).setCellValue(i * 10);
            }
            
            try (FileOutputStream out = new FileOutputStream(tempFile)) {
                wb.write(out);
            }
        }
        
        // Compare lazy vs regular loading
        try (XSSFWorkbook regularWb = new XSSFWorkbook(tempFile);
             LazyXSSFWorkbook lazyWb = new LazyXSSFWorkbook(tempFile)) {
            
            // Both should have same number of sheets
            assertEquals(regularWb.getNumberOfSheets(), lazyWb.getNumberOfSheets());
            
            // Sheet names should match
            for (int i = 0; i < regularWb.getNumberOfSheets(); i++) {
                assertEquals(regularWb.getSheetName(i), lazyWb.getSheetName(i));
            }
            
            // Data should match when accessed
            Sheet regularSheet = regularWb.getSheetAt(0);
            Sheet lazySheet = lazyWb.getSheetAt(0);
            
            assertEquals(regularSheet.getSheetName(), lazySheet.getSheetName());
            assertEquals(regularSheet.getLastRowNum(), lazySheet.getLastRowNum());
            
            // Check a few data points
            for (int i = 0; i < 5; i++) {
                Row regularRow = regularSheet.getRow(i);
                Row lazyRow = lazySheet.getRow(i);
                
                assertEquals(regularRow.getCell(0).getStringCellValue(), 
                           lazyRow.getCell(0).getStringCellValue());
                assertEquals(regularRow.getCell(1).getNumericCellValue(), 
                           lazyRow.getCell(1).getNumericCellValue());
            }
        }
    }
}