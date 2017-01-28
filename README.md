Excella PdfExporter
===============

1. ExCella PdfExporterとは  
---------------------

  ExCella PdfExporterはExcella Reportsで出力する帳票をPDF形式で出力するためのプログラムです。
  Excella Reports 1.6まではExcella Reportsに含んでいたものを、別プロジェクトとして分離しました。
  
  ExCella Reports、Excella PdfExporterは株式会社ビーブレイクシステムズが
  作成したオープンソースソフトウェアです。


2. 配布条件  
-------------

  本ソフトウェアはLGPL v3にて公開しています。
  詳細は「COPYING.LESSER」ファイルまたは
  以下のページを参照してください。
  http://www.gnu.org/licenses/lgpl-3.0-standalone.html


3. 免責  
---------

  このソフトウェアを使用したことによって生じたいかなる
  障害・損害・不具合等に関しても、株式会社ビーブレイクシステムズは
  一切の責任を負いません。各自の責任においてご使用ください。

4. Maven repository
-------------
mavenの依存ライブラリとして追加する場合、pom.xmlに下記のリポジトリを追加してください。
```xml
  <repositories>
    <repository>
	  <id>excella.bbreak.org</id>
      <name>bBreak Systems Excella</name>
      <url>http://excella-core.github.io/maven2/</url>    
    </repository>
  </repositories>
```

5. 使い方
-------------
PDF変換にはLibre Officeを使用します。
Libre Officeをインストール後、下記コマンドでサービスとして起動してください。

${LIBREOFFICE_HOME}/program/soffice -headless -accept=socket,port=8100;urp;

プログラム側では、ReportBookの出力形式としてOoPdfExporter.FORMAT_TYPEを指定します。
```
  ReportBook outputBook = new ReportBook( templateFilePath, outputFilePath, OoPdfExporter.FORMAT_TYPE);
```

出力時にOoPdfExporterを生成し、ReportProcessorに追加します。
```
  ReportProcessor reportProcessor = new ReportProcessor();
  OfficeManager officeManager =
    new ExternalOfficeManagerConfiguration().setPortNumber( 8100).buildOfficeManager();
  reportProcessor.addReportBookExporter( new OoPdfExporter(officeManager));
  reportProcessor.process( outputBook);
```

6. 更新履歴  
-------------
* 2017/01/28 Version 1.2 リリース
* 2016/04/05 Version 1.1 リリース
* 2016/01/17 Version 1.0 リリース


Copyright 2015 by bBreak Systems.
