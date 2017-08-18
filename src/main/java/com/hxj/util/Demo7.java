package com.hxj.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by Administrator on 2017/8/18.
 */
public class Demo7 {

        public static void main(String[] args)
        {
            HSSFWorkbook workbook = new HSSFWorkbook();


            HSSFSheet sheet = workbook.createSheet();
            HSSFRow row = sheet.createRow(1);


            System.out.println("创建单元格前的状态:");
            System.out.println("First:" + row.getFirstCellNum());
            System.out.println("Last:" + row.getLastCellNum());
            System.out.println("Total:" + row.getPhysicalNumberOfCells() + "\n");


            row.createCell((short) 0);


            System.out.println("创建第二列（列号为1）单元格:");
            System.out.println("First:" + row.getFirstCellNum());
            System.out.println("Last:" + row.getLastCellNum());
            System.out.println("Total:" + row.getPhysicalNumberOfCells() + "\n");


            row.createCell((short) 3);


            System.out.println("创建第四列（列号为3）单元格:");
            System.out.println("First:" + row.getFirstCellNum());
            System.out.println("Last:" + row.getLastCellNum());
            System.out.println("Total:" + row.getPhysicalNumberOfCells() + "\n");
        }
}
