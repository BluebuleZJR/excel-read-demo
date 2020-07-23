package org.jerryzeng.excel;

import java.io.IOException;
import java.io.InputStream;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.poifs.filesystem.FileMagic;

/**
 * @author JerryZeng
 * @date 2020/6/30
 */

public class ExcelFileCreateFactory {

  public static ExcelFile create(InputStream inputStream, String filename) throws IOException{
    filename = filename.toLowerCase();
    String extension = FilenameUtils.getExtension(filename);

    switch (extension) {
      case ExcelFile.FILE_XLS:
      case ExcelFile.FILE_XLSX:
        InputStream is = FileMagic.prepareToCheckMagic(inputStream);
        FileMagic fm = FileMagic.valueOf(is);
        switch (fm) {
          case OLE2:
            return new XlsFile(is);
          case OOXML:
            return new XlsxFile(is);
          default:
            throw new IOException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
        }
      case ExcelFile.FILE_CSV: return new CsvFile(inputStream);
      default:
        throw new IllegalArgumentException("Unknown file type");
    }
  }

}
