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
public class Demo2 {

    public static void main(String[] args)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();


        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(1);//创建第二行


        HSSFCell cell = row.createCell((short) 2);//创建第二行第三格
        cell.setCellValue("test");//第二行第三格写入test


        for (int i = 0; i < 3; i++)
        {
            HSSFCell c = row.getCell((short) i);
            if (c == null)
            {
                System.out.println("第" + i + "列单元格不存在");
            }
            else
            {
                System.out.println("第" + i + "列单元格获取成功");
            }
        }
    }
}
