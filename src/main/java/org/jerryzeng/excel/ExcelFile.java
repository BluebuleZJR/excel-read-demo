package org.jerryzeng.excel;

import java.io.Closeable;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * @author JerryZeng
 * @date 2020/6/30
 */
public interface ExcelFile extends Closeable {

  String FILE_XLSX = "xlsx";
  String FILE_XLS = "xls";
  String FILE_CSV = "csv";
  List<String> ALLOW_FILE_EXTENSIONS = Arrays.asList(FILE_XLSX, FILE_XLS, FILE_CSV);

  /**
   * 是否为空文件
   * @return true: 一行数据都没有或者只有表头，false: 至少有除表头外的一行数据
   * */
  boolean isEmpty();
  /**
   * 获取最后一行行号
   * @return 可能会为null
   * */
  int getLastRowNum();
  /**
   * 文件头信息，第一列的序号为0
   * @return head
   * */
  Map<Integer, String> getHead();
  /**
   * 获取指定表头和行数区间(闭区间)的值，例如表头为"手机号"的，第1行到第3行的数据即除表头外的前三行数据，
   * 第0行默认是表头
   * @param head 表头
   * @param from 起始行数
   * @param to 结束行数
   * @return 这个范围内单元格中值的集合
   * */
  List<String> getRangeCellValue(String head, int from, int to);

  /**
   * 获取文件的所有数据
   * @return 每行数据
   * */
  List<Map<String, String>> getFileData();

}
