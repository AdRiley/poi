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

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.Beta;

/**
 * Lazy-loading version of XSSFWorkbook that provides a memory-efficient approach
 * to working with large Excel files.
 * 
 * For now, this is a simple alias for XSSFWorkbook. In a full implementation,
 * this would defer loading of sheet data until actually needed, but implementing
 * true lazy loading requires significant changes to the POI architecture that
 * are beyond the scope of this immediate need.
 * 
 * The main benefit for InPlaceSXSSFWorkbook is that it signals the intent to
 * work with large files in a memory-efficient manner, and can be enhanced
 * in the future with true lazy loading capabilities.
 */
@Beta
public class LazyXSSFWorkbook extends XSSFWorkbook {

    /**
     * Constructs a LazyXSSFWorkbook from an existing Excel file.
     * 
     * @param file the Excel file to open
     * @throws IOException if an error occurs while reading the file
     * @throws InvalidFormatException if the file format is invalid
     */
    public LazyXSSFWorkbook(File file) throws IOException, InvalidFormatException {
        super(file);
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
     * Check if a sheet at the given index has been fully loaded.
     * For this simple implementation, all sheets are considered loaded.
     */
    public boolean isSheetLoaded(int index) {
        return index >= 0 && index < getNumberOfSheets();
    }

    /**
     * Get the number of fully loaded sheets.
     * For this simple implementation, this equals the total number of sheets.
     */
    public int getNumberOfLoadedSheets() {
        return getNumberOfSheets();
    }
}