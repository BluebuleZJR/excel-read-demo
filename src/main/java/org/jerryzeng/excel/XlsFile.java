package org.jerryzeng.excel;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

/**
 * @author JerryZeng
 * @date 2020/6/30
 */

public class XlsFile extends AbstractOfficeFile {

  private final Iterator<Row> iterator;

  public XlsFile(InputStream inputStream) throws IOException {
    super.workbook = new HSSFWorkbook(inputStream);
    super.defaultSheet = workbook.getSheetAt(0);
    this.iterator = super.defaultSheet.rowIterator();
    readHead();
  }

  @Override
  public Iterator<Row> iterator() {
    return this.iterator;
  }
}

