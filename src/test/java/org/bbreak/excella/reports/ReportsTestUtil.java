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

package org.bbreak.excella.reports;

import java.io.File;

/**
 * テスト用のユーティリティクラス
 * 
 * @since 1.0
 */
public class ReportsTestUtil {

    public static String getTestOutputDir() {

        String tempDir = System.getProperty( "user.dir") + "/work/test/";
        File file = new File( tempDir);
        if ( !file.exists()) {
            file.mkdirs();
        }

        return tempDir;
    }

}
