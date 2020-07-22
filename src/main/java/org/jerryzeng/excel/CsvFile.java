package org.jerryzeng.excel;

import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

/**
 * @author JerryZeng
 * @date 2020/7/3
 */

public class CsvFile extends AbstractExcelFile<String[]>{

  private final CSVReader reader;
  private final Iterator<String[]> iterator;

  public CsvFile(InputStream inputStream) throws IOException {
    BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(inputStream));
    reader = new CSVReaderBuilder(bufferedReader).build();
    iterator = reader.iterator();
    readHead();
  }

  @Override
  public int getLastRowNum() {
    throw new UnsupportedOperationException();
  }

  @Override
  public Iterator<String[]> iterator() {
    return iterator;
  }

  @Override
  protected Map<Integer, String> lineToHead(String[] first) {
    Map<Integer, String> head = new LinkedHashMap<>(first.length);
    for (int i = 0; i < first.length; i++) {
      head.put(i, first[i]);
    }
    return head;
  }

  @Override
  protected Map<String, String> lineToMap(String[] line) {
    Map<String, String> res = new HashMap<>(line.length);
    for (int i = 0; i < line.length; i++) {
      String headValue = head.get(i);
      if(headValue == null) {continue;}
      res.put(headValue, line[i]);
    }
    return res;
  }

  @Override
  public void close() throws IOException{
    reader.close();
  }
}
