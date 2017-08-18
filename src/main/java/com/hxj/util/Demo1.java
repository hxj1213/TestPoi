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
public class Demo1 {

    public static void main(String[] args)
    {
        HSSFWorkbook workbook = new HSSFWorkbook();


        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(1);//创建序号为1的行，第2行


        HSSFCell cell = row.createCell(2);//创建序号为2的单元格，第二行第3格
        cell.setCellValue("test");//写入test


        FileOutputStream out = null;
        try
        {
            out = new FileOutputStream("sample.xls");
            workbook.write(out);
        }
        catch (IOException e)
        {
            System.out.println(e.toString());
        }
        finally
        {
            try
            {
                out.close();
            }
            catch (IOException e)
            {
                System.out.println(e.toString());
            }
        }
    }

}
