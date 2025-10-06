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

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.Beta;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;

/**
 * Lazy-loading version of XSSFWorkbook that defers loading sheet data until needed.
 * 
 * This class extends XSSFWorkbook but overrides the sheet loading mechanism to only
 * load sheet metadata initially. The actual sheet data is loaded on-demand when
 * a sheet is first accessed.
 * 
 * This is particularly useful for large workbooks where you only need to work with
 * a subset of sheets, as it significantly reduces initial memory consumption and
 * load time.
 */
@Beta
public class LazyXSSFWorkbook extends XSSFWorkbook {
    
    /**
     * Map to store sheet metadata without loading the actual sheet data
     */
    private final Map<String, CTSheet> sheetMetadata = new HashMap<>();
    
    /**
     * Map to store the relationship between sheet IDs and XSSFSheet instances
     * This is used to track which sheets have been loaded
     */
    private final Map<String, XSSFSheet> loadedSheets = new HashMap<>();
    
    /**
     * Store the original sheet ID map for lazy loading
     */
    private Map<String, XSSFSheet> originalSheetIdMap;

    /**
     * Constructs a LazyXSSFWorkbook from an existing Excel file.
     * 
     * @param file the Excel file to open
     * @throws IOException if an error occurs while reading the file
     * @throws InvalidFormatException if the file format is invalid
     */
    public LazyXSSFWorkbook(File file) throws IOException, InvalidFormatException {
        super(OPCPackage.open(file));
    }

    /**
     * Constructs a LazyXSSFWorkbook from an OPC package.
     * 
     * @param pkg the OPC package
     * @throws IOException if an error occurs while reading the package
     */
    public LazyXSSFWorkbook(OPCPackage pkg) throws IOException {
        super(pkg);
    }

    /**
     * Override parseSheet to defer actual sheet loading.
     * We only store the sheet metadata here and load the sheet content on demand.
     */
    @Override
    public void parseSheet(Map<String, XSSFSheet> shIdMap, CTSheet ctSheet) {
        // Store the original sheet ID map for later use
        if (originalSheetIdMap == null) {
            originalSheetIdMap = new HashMap<>(shIdMap);
        }
        
        // Instead of loading the sheet immediately, just store the metadata
        sheetMetadata.put(ctSheet.getId(), ctSheet);
        
        // Create a placeholder in the sheets list to maintain proper indexing
        // We'll replace this with the actual sheet when it's first accessed
        sheets.add(null);
    }

    /**
     * Override getSheetAt to implement lazy loading.
     * The sheet is only loaded when first accessed.
     */
    @Override
    public XSSFSheet getSheetAt(int index) {
        if (index < 0 || index >= sheets.size()) {
            throw new IllegalArgumentException("Sheet index (" + index + ") is out of range (0.." + (sheets.size()-1) + ")");
        }
        
        XSSFSheet sheet = sheets.get(index);
        if (sheet == null) {
            // Sheet hasn't been loaded yet, load it now
            sheet = loadSheetAtIndex(index);
            sheets.set(index, sheet);
        }
        
        return sheet;
    }

    /**
     * Override getSheet to implement lazy loading by name.
     */
    @Override
    public XSSFSheet getSheet(String name) {
        for (int i = 0; i < sheets.size(); i++) {
            XSSFSheet sheet = sheets.get(i);
            if (sheet == null) {
                // Check if this unloaded sheet has the name we're looking for
                CTSheet ctSheet = getSheetMetadataByIndex(i);
                if (ctSheet != null && name.equals(ctSheet.getName())) {
                    sheet = loadSheetAtIndex(i);
                    sheets.set(i, sheet);
                    return sheet;
                }
            } else if (name.equals(sheet.getSheetName())) {
                return sheet;
            }
        }
        return null;
    }

    /**
     * Override getSheetName to avoid loading the sheet just to get its name.
     */
    @Override
    public String getSheetName(int sheet) {
        CTSheet ctSheet = getSheetMetadataByIndex(sheet);
        return ctSheet != null ? ctSheet.getName() : super.getSheetName(sheet);
    }

    /**
     * Override iterator to ensure lazy loading works with iteration.
     */
    @Override
    public java.util.Iterator<Sheet> iterator() {
        return new LazySheetIterator();
    }

    /**
     * Load a sheet at the specified index on demand.
     */
    private XSSFSheet loadSheetAtIndex(int index) {
        CTSheet ctSheet = getSheetMetadataByIndex(index);
        if (ctSheet == null) {
            throw new IllegalStateException("No sheet metadata found for index " + index);
        }
        
        XSSFSheet xssfSheet = originalSheetIdMap.get(ctSheet.getId());
        if (xssfSheet == null) {
            throw new IllegalStateException("No sheet found for ID " + ctSheet.getId());
        }
        
        // Now load the sheet data
        xssfSheet.sheet = ctSheet;
        xssfSheet.onDocumentRead();
        
        loadedSheets.put(ctSheet.getId(), xssfSheet);
        return xssfSheet;
    }

    /**
     * Get sheet metadata by index without loading the sheet.
     */
    private CTSheet getSheetMetadataByIndex(int index) {
        if (workbook != null && workbook.getSheets() != null && workbook.getSheets().getSheetArray() != null) {
            CTSheet[] sheetArray = workbook.getSheets().getSheetArray();
            if (index >= 0 && index < sheetArray.length) {
                return sheetArray[index];
            }
        }
        return null;
    }

    /**
     * Check if a sheet at the given index has been loaded.
     */
    public boolean isSheetLoaded(int index) {
        return index >= 0 && index < sheets.size() && sheets.get(index) != null;
    }

    /**
     * Get the number of loaded sheets.
     */
    public int getNumberOfLoadedSheets() {
        return loadedSheets.size();
    }

    /**
     * Iterator that implements lazy loading of sheets.
     */
    private class LazySheetIterator implements java.util.Iterator<Sheet> {
        private int index = 0;
        
        @Override
        public boolean hasNext() {
            return index < getNumberOfSheets();
        }
        
        @Override
        public Sheet next() {
            return getSheetAt(index++);
        }
        
        @Override
        public void remove() {
            throw new UnsupportedOperationException("Remove not supported");
        }
    }
}