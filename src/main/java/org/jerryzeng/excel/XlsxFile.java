package org.jerryzeng.excel;

import com.monitorjbl.xlsx.StreamingReader;
import java.io.InputStream;

/**
 * @author JerryZeng
 * @date 2020/6/30
 */
public class XlsxFile extends AbstractOfficeFile {

  public XlsxFile(InputStream inputStream) {
    super.workbook = StreamingReader.builder()
      .rowCacheSize(100)
      .bufferSize(4096)
      .open(inputStream);
    super.defaultSheet = workbook.getSheetAt(0);
    readHead();
  }

}
