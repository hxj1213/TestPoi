package com.hxj.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

/**
 * Created by Administrator on 2017/8/18.
 */
public class Demo4
{
    public static void main(String[] args){
        HSSFWorkbook workbook = new HSSFWorkbook();


        HSSFSheet sheet = workbook.createSheet();
        HSSFRow row = sheet.createRow(1);//创建第二行


        HSSFCell cell1 = row.createCell((short) 0);//2,1格
        cell1.setCellValue(true);//写入true


        HSSFCell cell2 = row.createCell((short) 1);//2,2格
        Calendar cal = Calendar.getInstance();//Calendar？？？
        cell2.setCellValue(cal);//写入Calendar型对象cal


        HSSFCell cell3 = row.createCell((short) 2);//2,3格
        Date date = new Date(); //日期型
        cell3.setCellValue(date);//写入日期型


        HSSFCell cell4 = row.createCell((short) 3);//2,4格
        cell4.setCellValue(150);//写入150


        HSSFCell cell5 = row.createCell((short) 4);//2.5格
        cell5.setCellValue("hello");//写入hello


        HSSFRow row2 = sheet.createRow(2);//第三行


        HSSFCell cell6 = row2.createCell((short) 0);//3,1格


        FileOutputStream out = null;
        try
        {
            out = new FileOutputStream("sample1.xls");
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
