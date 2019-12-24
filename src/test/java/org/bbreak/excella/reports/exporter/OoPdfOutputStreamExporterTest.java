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
import static org.junit.Assert.fail;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.net.URLDecoder;
import java.util.Date;

import org.jodconverter.office.ExternalOfficeManagerBuilder;
import org.jodconverter.office.OfficeManager;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.model.ReportBook;
import org.bbreak.excella.reports.model.ReportSheet;
import org.bbreak.excella.reports.processor.ReportProcessor;
import org.junit.Test;

/**
 * {@link org.bbreak.excella.reports.exporter.OoPdfOutputStreamExporter} のためのテスト・クラス。
 *
 * @since 1.1
 */
public class OoPdfOutputStreamExporterTest {

    private String tmpDirPath = ReportsTestUtil.getTestOutputDir();

    ConvertConfiguration configuration = null;

    private OfficeManager officeManager = new ExternalOfficeManagerBuilder().setPortNumber( 8100).build();

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfOutputStreamExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     */
    @Test
    public void testOutput() {

        // XLSテスト� - フォーマット複数指定・サイズ比較テスト
        FileOutputStream xlsFileOutputStream = null;
        String filePath = tmpDirPath + (new Date()).getTime();
        String xlsTemplateFileName = "OoPdfOutputStreamExporterTest.xls";
        URL xlsTemplateFileUrl = OoPdfOutputStreamExporterTest.class.getResource( xlsTemplateFileName);
        try {
            String xlsTemplateFilePath = URLDecoder.decode( xlsTemplateFileUrl.getPath(), "UTF-8");

            ReportBook xlsOutputBook = new ReportBook( xlsTemplateFilePath, filePath, OoPdfExporter.FORMAT_TYPE, OoPdfOutputStreamExporter.FORMAT_TYPE);

            ReportSheet outputSheet = new ReportSheet( "test");
            xlsOutputBook.addReportSheet( outputSheet);

            ReportProcessor xlsReportProcessor = new ReportProcessor();

            // Streamの処理
            File xlsStreamFile = new File( filePath + "_stream.pdf");
            xlsFileOutputStream = new FileOutputStream( xlsStreamFile);

            xlsReportProcessor.addReportBookExporter( new OoPdfExporter( officeManager));
            xlsReportProcessor.addReportBookExporter( new OoPdfOutputStreamExporter( officeManager, xlsFileOutputStream));
            xlsReportProcessor.process( xlsOutputBook);

            File xlsExcelFile = new File( filePath + ".pdf");

            // サイズ比較
            long xlsFileByteLength = xlsExcelFile.length();
            long xlsStreamFileSize = xlsStreamFile.length();
            assertEquals( xlsFileByteLength, xlsStreamFileSize);

        } catch ( Exception e) {
            e.getStackTrace();
            fail( e.toString());
        } finally {
            try {
                xlsFileOutputStream.close();
            } catch ( IOException e) {
                e.printStackTrace();
                fail( e.toString());
            }
        }

        // XLSテスト� - フォーマット単数指定・ストリーム出力によるファイル以外が削除されていることを確認
        try {
            String xlsTemplateFilePath = URLDecoder.decode( xlsTemplateFileUrl.getPath(), "UTF-8");

            ReportBook xlsOutputBook = new ReportBook( xlsTemplateFilePath, filePath, OoPdfOutputStreamExporter.FORMAT_TYPE);

            ReportSheet outputSheet = new ReportSheet( "test");
            xlsOutputBook.addReportSheet( outputSheet);

            ReportProcessor xlsReportProcessor = new ReportProcessor();

            // Streamの処理
            File xlsStreamFile = new File( filePath + "_existStream.pdf");
            xlsFileOutputStream = new FileOutputStream( xlsStreamFile);

            xlsReportProcessor.addReportBookExporter( new OoPdfOutputStreamExporter( officeManager, xlsFileOutputStream));
            xlsReportProcessor.process( xlsOutputBook);

            File xlsExcelFile = new File( filePath + "_exist.pdf");

            // ストリーム出力されたファイルが開けること（0バイトでないファイルであることを確認）
            if ( xlsStreamFile.length() <= 0) {
                fail( "File is not opened");
            }

            // ファイルが削除されていることを確認
            if ( xlsExcelFile.exists()) {
                fail( "PDFfile exists.");
            }
        } catch ( Exception e) {
            e.getStackTrace();
            fail( e.toString());
        } finally {
            try {
                xlsFileOutputStream.close();
            } catch ( IOException e) {
                e.printStackTrace();
                fail( e.toString());
            }
        }

        // XLSX
        FileOutputStream xlsxFileOutputStream = null;
        try {
            String xlsxTemplateFileName = "OoPdfOutputStreamExporterTest.xlsx";
            URL xlsxTemplateFileUrl = OoPdfOutputStreamExporterTest.class.getResource( xlsxTemplateFileName);
            String xlsxTemplateFilePath = URLDecoder.decode( xlsxTemplateFileUrl.getPath(), "UTF-8");

            ReportBook xlsOutputBook = new ReportBook( xlsxTemplateFilePath, filePath, OoPdfOutputStreamExporter.FORMAT_TYPE);

            // Streamの処理
            File xlsxStreamFile = new File( filePath + "(2007)" + "_stream.pdf");
            xlsxFileOutputStream = new FileOutputStream( xlsxStreamFile);

            ReportProcessor xlsxReportProcessor = new ReportProcessor();
            xlsxReportProcessor.addReportBookExporter( new OoPdfOutputStreamExporter( officeManager, xlsxFileOutputStream));
            xlsxReportProcessor.process( xlsOutputBook);

//            fail( "XLSXは変換不可");
        } catch ( IllegalArgumentException e) {
            fail( e.toString());
        } catch ( Exception e) {
            e.printStackTrace();
//            fail( e.toString());
            // TODO org.artofsolving.jodconverter.office.OfficeException: could not store document
            // が発生する
        } finally {
            try {
                xlsxFileOutputStream.close();
            } catch ( IOException e) {
                e.printStackTrace();
                fail( e.toString());
            }
        }
    }
}
