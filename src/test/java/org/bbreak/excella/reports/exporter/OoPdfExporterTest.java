/*-
 * #%L
 * excella-pdfexporter
 * %%
 * Copyright (C) 2009 - 2019 bBreak Systems and contributors
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *      http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * #L%
 */

package org.bbreak.excella.reports.exporter;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import java.io.File;
import java.util.Date;

import org.apache.poi.ss.usermodel.Workbook;
import org.jodconverter.office.ExternalOfficeManagerBuilder;
import org.jodconverter.office.OfficeManager;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.exporter.OoPdfExporter} のためのテスト・クラス。
 * 
 * @since 1.0
 */
public class OoPdfExporterTest extends ReportsWorkbookTest {

    public OoPdfExporterTest( String version) {
        super( version);
    }

    private String tmpDirPath = ReportsTestUtil.getTestOutputDir();

    ConvertConfiguration configuration = null;
    
    private OfficeManager officeManager = new ExternalOfficeManagerBuilder().setPortNumber( 8100).build();

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     */
    @Test
    public void testOutput() {

        OoPdfExporter exporter = new OoPdfExporter( officeManager);
        String filePath = null;

        Workbook wb = getWorkbook();

            wb = getWorkbook();
            configuration = new ConvertConfiguration( OoPdfExporter.EXTENTION);
            filePath = tmpDirPath + System.currentTimeMillis() + exporter.getExtention();
            exporter.setFilePath( filePath);
            try {
                exporter.output( wb, new BookData(), configuration);
                File file = new File( exporter.getFilePath());
                assertTrue( file.exists());
            } catch ( ExportException e) {
                e.printStackTrace();
                fail( e.toString());
            }

            // オプション指定
            wb = getWorkbook();
            configuration.addOption( "PermissionPassword", "pass");
            configuration.addOption( "RestrictPermissions", Boolean.TRUE);
            configuration.addOption( "Printing", 0);
            configuration.addOption( "Changes", 4);
            filePath = tmpDirPath + System.currentTimeMillis() + exporter.getExtention();
            exporter.setFilePath( filePath);
            try {
                exporter.output( wb, new BookData(), configuration);
                File file = new File( exporter.getFilePath());
                assertTrue( file.exists());
            } catch ( ExportException e) {
                e.printStackTrace();
                fail( e.toString());
            }

            // 例外発生
            wb = getWorkbook();
            configuration = new ConvertConfiguration( OoPdfExporter.EXTENTION);
            filePath = tmpDirPath + (new Date()).getTime() + exporter.getExtention();
            exporter.setFilePath( filePath);
            try {
                exporter.output( wb, new BookData(), configuration);
            } catch ( ExportException e) {
                fail( e.toString());
            }
            File file = new File( exporter.getFilePath());
            file.setReadOnly();
            try {
                exporter.output( wb, new BookData(), configuration);
                fail( "例外未発生");
            } catch ( Exception e) {
                if ( e instanceof ExportException) {
                    // OK
                } else {
                    fail( e.toString());
                }
            }

    }

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#getFormatType()} のためのテスト・メソッド。
     */
    @Test
    public void testGetFormatType() {
        OoPdfExporter exporter = new OoPdfExporter( officeManager);
        assertEquals( "PDF", exporter.getFormatType());
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#getExtention()} のためのテスト・メソッド。
     */
    @Test
    public void testGetExtention() {
        OoPdfExporter exporter = new OoPdfExporter( officeManager);
        assertEquals( ".pdf", exporter.getExtention());
    }

}
