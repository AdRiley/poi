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

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Streaming version of XSSFWorkbook that works with existing files in-place.
 * 
 * This class extends SXSSFWorkbook to provide the ability to open an existing
 * Excel file and stream modifications to it, rather than starting with a blank
 * workbook or template.
 * 
 * The class loads an existing XLSX file and creates a streaming workbook from it,
 * allowing for memory-efficient operations on large existing files.
 */
public class InPlaceSXSSFWorkbook extends SXSSFWorkbook {

    /**
     * Constructs an InPlaceSXSSFWorkbook from an existing Excel file.
     * 
     * @param file the Excel file to open and stream
     * @throws IOException if an error occurs while reading the file
     * @throws InvalidFormatException if the file format is invalid
     */
    public InPlaceSXSSFWorkbook(File file) throws IOException, InvalidFormatException {
        super(new XSSFWorkbook(file));
    }

    /**
     * Constructs an InPlaceSXSSFWorkbook from an existing Excel file with a specified row access window size.
     * 
     * @param file the Excel file to open and stream
     * @param rowAccessWindowSize the number of rows that are kept in memory until flushed out
     * @throws IOException if an error occurs while reading the file
     * @throws InvalidFormatException if the file format is invalid
     */
    public InPlaceSXSSFWorkbook(File file, int rowAccessWindowSize) throws IOException, InvalidFormatException {
        super(new XSSFWorkbook(file), rowAccessWindowSize);
    }

    /**
     * Constructs an InPlaceSXSSFWorkbook from an existing Excel file with specified options.
     * 
     * @param file the Excel file to open and stream
     * @param rowAccessWindowSize the number of rows that are kept in memory until flushed out
     * @param compressTmpFiles whether to use gzip compression for temporary files
     * @throws IOException if an error occurs while reading the file
     * @throws InvalidFormatException if the file format is invalid
     */
    public InPlaceSXSSFWorkbook(File file, int rowAccessWindowSize, boolean compressTmpFiles) throws IOException, InvalidFormatException {
        super(new XSSFWorkbook(file), rowAccessWindowSize, compressTmpFiles);
    }

    /**
     * Constructs an InPlaceSXSSFWorkbook from an existing Excel file with all options.
     * 
     * @param file the Excel file to open and stream
     * @param rowAccessWindowSize the number of rows that are kept in memory until flushed out
     * @param compressTmpFiles whether to use gzip compression for temporary files
     * @param useSharedStringsTable whether to use a shared strings table
     * @throws IOException if an error occurs while reading the file
     * @throws InvalidFormatException if the file format is invalid
     */
    public InPlaceSXSSFWorkbook(File file, int rowAccessWindowSize, boolean compressTmpFiles, boolean useSharedStringsTable) throws IOException, InvalidFormatException {
        super(new XSSFWorkbook(file), rowAccessWindowSize, compressTmpFiles, useSharedStringsTable);
    }
}