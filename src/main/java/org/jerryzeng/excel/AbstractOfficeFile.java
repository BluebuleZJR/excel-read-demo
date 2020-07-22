package org.jerryzeng.excel;

import java.io.IOException;
import java.text.DecimalFormat;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author JerryZeng
 * @date 2020/6/30
 */

public abstract class AbstractOfficeFile extends AbstractExcelFile<Row> {

  protected Workbook workbook;
  protected Sheet defaultSheet;

  @Override
  public int getLastRowNum() {
    return defaultSheet.getLastRowNum();
  }

  @Override
  public void close() throws IOException{
      workbook.close();
  }

  @Override
  protected Map<Integer, String> lineToHead(Row first) {
    Map<Integer, String> head = new LinkedHashMap<>(16);
    for(Cell cell : first) {
      String value = getStringCellValue(cell);
      head.put(cell.getColumnIndex(), value);
    }
    return head;
  }

  @Override
  protected Map<String, String> lineToMap(Row line) {
    Map<String, String> res = new HashMap<>(16);
    for(Cell cell : line) {
      String keyValue = head.get(cell.getColumnIndex());
      if(keyValue == null) {continue;}
      res.put(keyValue, getStringCellValue(cell));
    }
    return res;
  }

  @Override
  public Iterator<Row> iterator() {
    return defaultSheet.iterator();
  }

  private String getStringCellValue(Cell cell) {
    if(cell == null) {return "";}

    switch (cell.getCellType()) {
      case NUMERIC:
        if (DateUtil.isCellDateFormatted(cell)) {
          return cell.getDateCellValue().toString();
        } else {
          return new DecimalFormat("#").format(cell.getNumericCellValue());
        }
      case STRING:
        return String.valueOf(cell.getStringCellValue());
      case BOOLEAN:
        return String.valueOf(cell.getBooleanCellValue());
      case FORMULA:
        return String.valueOf(cell.getCellFormula());
      case BLANK:
        return cell.getStringCellValue();
      default:
        return  "";
    }
  }
}
