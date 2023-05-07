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
import java.io.IOException;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Date;

import org.apache.poi.ss.usermodel.Workbook;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.ReportsTestUtil;
import org.bbreak.excella.reports.model.ConvertConfiguration;
import org.bbreak.excella.reports.processor.ReportsWorkbookTest;
import org.jodconverter.core.office.OfficeException;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.office.LocalOfficeManager;
import org.junit.After;
import org.junit.Before;
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

    private OfficeManager officeManager = LocalOfficeManager.builder().portNumbers( 8100).build();

    @Before
    public void startOfficeManager() throws OfficeException {
        officeManager.start();
    }

    @After
    public void stopOfficeManager() throws OfficeException {
        officeManager.stop();
    }

    /**
     * {@link org.bbreak.excella.reports.exporter.OoPdfExporter#output(org.apache.poi.ss.usermodel.Workbook, org.bbreak.excella.core.BookData, org.bbreak.excella.reports.model.ConvertConfiguration)}
     * のためのテスト・メソッド。
     * 
     * @throws OfficeException
     * @throws IOException (CI only) 外部プロセスの起動に失敗した場合
     * @throws InterruptedException (CI only) 外部プロセスが規定時間内に完了しない場合
     */
    @Test
    public void testOutput() throws OfficeException, IOException, InterruptedException {

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
            String fileName = new Date().getTime() + exporter.getExtention();
            filePath = tmpDirPath + fileName;
            if ( System.getProperty( "os.name").toLowerCase().contains( "nux")) {
                // Linux環境ではsetReadonlyしてもLibreOfficeがファイルを上書きできてしまうため、おそらく書き込み許可がないであろうディレクトリに出力する
                // ※Ubuntu 22.04 + ext4 fs + LibreOffice 7.3.7.2で確認。手順は「新規作成→chmod 444したファイルを指定して上書き保存」
                // ods形式であればchmod 444したファイルを上書きできないが、xlsx形式の保存やpdfエクスポートでは上書きされてしまう
                exporter.setFilePath( "/proc/" + fileName);
                Runtime.getRuntime().addShutdownHook( new Thread( () -> cleanFile( exporter.getFilePath())));
            } else {
                // すでに存在するファイル(書き込み許可なし)の上書きを試みる
                exporter.setFilePath( filePath);
                try {
                    exporter.output( wb, new BookData(), configuration);
                } catch ( ExportException e) {
                    fail( e.toString());
                }
                File file = new File( exporter.getFilePath());
                file.setReadOnly();
            }

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

    private void cleanFile( String path) {
        try {
            Files.deleteIfExists( Paths.get( path));
        } catch ( IOException e) {
            throw new UncheckedIOException( e);
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
