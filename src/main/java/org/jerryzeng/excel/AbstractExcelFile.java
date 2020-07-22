package org.jerryzeng.excel;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;

/**
 * @author JerryZeng
 * @date 2020/7/3
 */
public abstract class AbstractExcelFile<T> implements ExcelFile, Iterable<T> {
  /**
   * 表头
   * key 是列数从零开始计数
   * value 是表头的值
   * 例如 {0=姓名,1=年龄}
   * */
  protected Map<Integer, String> head;
  /**
   * 文件内的数据
   * key 表头名
   * value 值
   * 例如 {姓名=jerry,年龄=25}
   * */
  protected List<Map<String, String>> data = new ArrayList<>();
  /**
   * 当前迭代到的行数，即已经读了多少行了，表头也算一行
   * */
  protected int currentIteratorIndex;

  protected void incrRowIndex() {
    currentIteratorIndex += 1;
  }

  @Override
  public boolean isEmpty() {
    if(currentIteratorIndex == 0) {
      readHead();
    }
    if(MapUtils.isEmpty(head)) {
      return true;
    }
    if(currentIteratorIndex == 1) {
      readNextDataLine();
    }

    return CollectionUtils.isEmpty(data);
  }

  @Override
  public Map<Integer, String> getHead() {
    if(currentIteratorIndex == 0) {
      readHead();
    }
    return head;
  }

  @Override
  public List<String> getRangeCellValue(String head, int from, int to) {
    if(isUnreadHead()) {
      throw new UnsupportedOperationException("No head can't go on");
    }
    if((from < 0 || to < 0) || (from > to)) {
      throw new IllegalArgumentException("from(" + from + ") > to(" + to + ")");
    }
    List<String> res = new ArrayList<>();
    if (iterator().hasNext()) {
      while (iterator().hasNext() && data.size() < to) {
        readNextDataLine();
      }
    }

    if(CollectionUtils.isEmpty(data) || data.size() < from) {
      while (iterator().hasNext() && data.size() < to) {
        readNextDataLine();
      }
    }

    if(CollectionUtils.isEmpty(data)) {return res;}
    if(data.size() < from) {return res;}

    for (int i = from-1; i <= to-1 && i < data.size(); i++) {
      Map<String, String> line = data.get(i);
      res.add(line.get(head));
    }
    return res;
  }

  @Override
  public List<Map<String, String>> getFileData() {
    while(iterator().hasNext()) {
      readNextDataLine();
    }

    return data;
  }

  protected void readHead() {
    if(head != null) {
      return;
    }

    if(currentIteratorIndex != 0) {
      throw new UnsupportedOperationException("表头已去读");
    }

    if(!iterator().hasNext()) {
      head = Collections.emptyMap();
      incrRowIndex();
      return;
    }

    head = lineToHead(iterator().next());
    incrRowIndex();
  }

  protected void readNextDataLine() {
    if(isUnreadHead()) {
      throw new UnsupportedOperationException("No head can't go on");
    }

    if(!iterator().hasNext()) {
      incrRowIndex();
      return;
    }

    data.add(lineToMap(iterator().next()));
    incrRowIndex();
  }

  protected abstract Map<Integer, String> lineToHead(T first);

  protected abstract Map<String, String> lineToMap(T line);

  protected boolean isUnreadHead() {
    return currentIteratorIndex == 0 || MapUtils.isEmpty(head);
  }
}
