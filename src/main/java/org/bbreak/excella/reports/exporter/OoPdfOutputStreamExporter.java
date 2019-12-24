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

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.jodconverter.office.OfficeManager;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.model.ConvertConfiguration;

/**
 * PDFをStreamに書き出すExporter
 *
 * @since 1.1
 */
public class OoPdfOutputStreamExporter extends OoPdfExporter {

    /**
     * 一時ファイルのプレフィックス
     */
    private static final String TMP_FILE_PREFIX = "tmp";

    /**
     * 変換タイプ：PDF
     */
    public static final String FORMAT_TYPE = "OUTPUT_STREAM_PDF";

    /**
     * 拡張子：PDF
     */
    public static final String EXTENTION = ".pdf";

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( OoPdfOutputStreamExporter.class);

    /**
     * 出力ストリーム
     */
    private OutputStream outputStream;

    /**
     * コンストラクタ
     *
     * @param outputStream 出力ストリーム
     */
    public OoPdfOutputStreamExporter( OutputStream outputStream) {
        this.outputStream = outputStream;
    }

    public OoPdfOutputStreamExporter( OfficeManager officeManager, OutputStream outputStream) {
        super( officeManager);
        this.outputStream = outputStream;
    }

    @Override
    public String getFormatType() {
        return FORMAT_TYPE;
    }

    @Override
    public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
        if ( log.isInfoEnabled()) {
            log.info( "処理結果を" + outputStream.getClass().getCanonicalName() + "に出力します");
        }

        // 一時的にファイル出力
        int point = getFilePath().indexOf( EXTENTION);
        StringBuffer sb = new StringBuffer( getFilePath());
        sb.insert( point, TMP_FILE_PREFIX);
        String tmpFilePath = sb.toString();
        setFilePath( tmpFilePath);

        // 親クラスでPDFの出力まで実行
        super.output( book, bookdata, configuration);

        File pdfFile = new File( getFilePath());

        // ファイルの内容をStreamに書き出す
        BufferedInputStream in = null;
        BufferedOutputStream out = null;
        try {
            in = new BufferedInputStream( new FileInputStream( pdfFile));
            out = new BufferedOutputStream( outputStream);
            int b;

            while ( (b = in.read()) != -1) {
                out.write( b);
            }
        } catch ( IOException e) {
            throw new ExportException( e);
        } finally {
            try {
                if ( in != null) {
                    in.close();
                }
                if ( out != null) {
                    out.close();
                }
            } catch ( IOException e) {
                throw new ExportException( e);
            } finally {
                // PDFファイルの削除
                pdfFile.delete();
            }
        }
    }
}
