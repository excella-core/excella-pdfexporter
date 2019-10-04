/*************************************************************************
 *
 * Copyright 2009 by bBreak Systems.
 *
 * ExCella Reports - Excelファイルを利用した帳票ツール
 *
 * $Id: OoPdfExporter.java 97 2010-01-13 02:11:36Z tomo-shibata $
 * $Revision: 97 $
 *
 * This file is part of ExCella Reports.
 *
 * ExCella Reports is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License version 3
 * only, as published by the Free Software Foundation.
 *
 * ExCella Reports is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License version 3 for more details
 * (a copy is included in the COPYING.LESSER file that accompanied this code).
 *
 * You should have received a copy of the GNU Lesser General Public License
 * version 3 along with ExCella Reports .  If not, see
 * <http://www.gnu.org/licenses/lgpl-3.0-standalone.html>
 * for a copy of the LGPLv3 License.
 *
 ************************************************************************/
package org.bbreak.excella.reports.exporter;

import java.io.File;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Workbook;
import org.jodconverter.OfficeDocumentConverter;
import org.jodconverter.document.DefaultDocumentFormatRegistry;
import org.jodconverter.document.DocumentFamily;
import org.jodconverter.document.DocumentFormat;
import org.jodconverter.document.DocumentFormatRegistry;
import org.jodconverter.document.SimpleDocumentFormatRegistry;
import org.jodconverter.office.DefaultOfficeManagerBuilder;
import org.jodconverter.office.OfficeException;
import org.jodconverter.office.OfficeManager;
import org.bbreak.excella.core.BookData;
import org.bbreak.excella.core.exception.ExportException;
import org.bbreak.excella.reports.model.ConvertConfiguration;

/**
 * OpenOfficePDF出力エクスポーター
 *
 * @since 1.0
 */
public class OoPdfExporter extends ReportBookExporter {

    /**
     * ログ
     */
    private static Log log = LogFactory.getLog( OoPdfExporter.class);

    /**
     * 変換タイプ：PDF
     */
    public static final String FORMAT_TYPE = "PDF";

    /**
     * 拡張子：PDF
     */
    public static final String EXTENTION = ".pdf";

    /**
     * OpneOfficeデフォルトポート番号
     */
    public static final int OPENOFFICE_DEFAULT_PORT = 8100;

    /**
     * OpneOfficeポート番号
     */
    private int port = OPENOFFICE_DEFAULT_PORT;

    /**
     * OpneOfficeマネージャ
     */
    private OfficeManager officeManager = null;

    /**
     * OpneOfficeマネージャのコントロール有無
     */
    private boolean controlOfficeManager = false;

    /**
     * コンストラクタ<BR>
     * デフォルトポート番号8100
     */
    public OoPdfExporter() {
    }

    /**
     * コンストラクタ
     *
     * @param port OpneOfficeポート番号
     */
    public OoPdfExporter( int port) {
        this.port = port;
    }

    /**
     * コンストラクタ
     *
     * @param officeManager OpneOfficeマネージャ
     */
    public OoPdfExporter( OfficeManager officeManager) {
        this.officeManager = officeManager;
        controlOfficeManager = true;
    }

    /*
     * (non-Javadoc)
     *
     * @see org.poireports.exporter.ReportBookExporter#output(org.apache.poi.ss.usermodel.Workbook, org.excelparser.BookData, org.poireports.model.ConvertConfiguration)
     */
    @Override
    public void output( Workbook book, BookData bookdata, ConvertConfiguration configuration) throws ExportException {
        // TODO POIにより出力されたxlsxファイルはOpenOfficeで読めない不具合あり
        // http://www.openoffice.org/issues/show_bug.cgi?id=97460
        // https://issues.apache.org/bugzilla/show_bug.cgi?id=46419
        // POI3.5より対応済
//        if ( book instanceof XSSFWorkbook) {
//            throw new IllegalArgumentException( "XSSFFile not supported.");
//        }

        if ( log.isInfoEnabled()) {
            log.info( "処理結果を" + getFilePath() + "に出力します");
        }

        if ( !controlOfficeManager) {
            officeManager = new DefaultOfficeManagerBuilder().setPortNumbers( port).build();
            try {
                officeManager.start();
            } catch ( OfficeException e) {
                throw new ExportException( e);
           }
        }

        File tmpFile = null;
        try {

            OfficeDocumentConverter converter = null;
            if ( configuration.getOptionsProperties().isEmpty()) {
                converter = new OfficeDocumentConverter( officeManager);
            } else {
                DocumentFormatRegistry registry = createDocumentFormatRegistry( configuration);
                converter = new OfficeDocumentConverter( officeManager, registry);
            }

            // 一時フォルダに吐き出し
            ExcelExporter excelExporter = new ExcelExporter();
            tmpFile = File.createTempFile( getClass().getSimpleName(), null);
            String tmpFileName = tmpFile.getCanonicalPath();
            excelExporter.setFilePath( tmpFileName);
            excelExporter.output( book, bookdata, null);

            tmpFileName = excelExporter.getFilePath();
            tmpFile = new File( tmpFileName);
            converter.convert( tmpFile, new File( getFilePath()));

        } catch ( Exception e) {
            throw new ExportException( e);
        } finally {

            if ( tmpFile != null) {
                // EXCEL削除
                tmpFile.delete();
            }
            if ( !controlOfficeManager) {
                try {
                    officeManager.stop();
                } catch ( OfficeException e) {
                    throw new ExportException( e);
                }
            }
        }
    }

    /**
     * 変換フォーマット情報を作成する。
     *
     * @param configuration 変換情報
     * @return 変換フォーマット情報
     */
    private DocumentFormatRegistry createDocumentFormatRegistry( ConvertConfiguration configuration) {

        SimpleDocumentFormatRegistry registry = ( SimpleDocumentFormatRegistry) DefaultDocumentFormatRegistry.getInstance();

        if ( configuration == null || configuration.getOptionsProperties().isEmpty()) {
            return registry;
        }

        DocumentFormat sourceFormat = registry.getFormatByExtension( "pdf");
        DocumentFormat modifiedFormat = DocumentFormat.builder() //
            .from( sourceFormat) //
            .storeProperty( DocumentFamily.SPREADSHEET, "FilterData", configuration.getOptions()) //
            .build();

        registry.addFormat( modifiedFormat);

        return registry;
    }

    /*
     * (non-Javadoc)
     *
     * @see org.poireports.exporter.ReportBookExporter#getFormatType()
     */
    @Override
    public String getFormatType() {
        return FORMAT_TYPE;
    }

    /*
     * (non-Javadoc)
     *
     * @see org.bbreak.excella.reports.exporter.ReportBookExporter#getExtention()
     */
    @Override
    public String getExtention() {
        return EXTENTION;
    }

}
