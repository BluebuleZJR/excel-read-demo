package org.jerryzeng.excel;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * @author JerryZeng
 * @date 2020/7/23
 */
public class Main {


  public static void main(String[] args) throws IOException {
//    FileInputStream is = new FileInputStream("/Users/jerryzeng/Downloads/test.xlsx");
//    ExcelFile file = ExcelFileCreateFactory.create(is, "test.xlsx");
//
//    System.out.println("head: " + file.getHead());
//    System.out.println(file.getLastRowNum());
//    file.getFileData().forEach(System.out::println);

//    FileInputStream is = new FileInputStream("/Users/jerryzeng/Downloads/test.xls");
//    ExcelFile file = ExcelFileCreateFactory.create(is, "test.xls");
//
//    System.out.println("head: " + file.getHead());
//    System.out.println(file.getLastRowNum());
//    file.getFileData().forEach(System.out::println);

    FileInputStream is = new FileInputStream("/Users/jerryzeng/Downloads/test.csv");
    ExcelFile file = ExcelFileCreateFactory.create(is, "test.csv");

    System.out.println("head: " + file.getHead());
    //System.out.println(file.getLastRowNum());
    file.getFileData().forEach(System.out::println);
  }
}
