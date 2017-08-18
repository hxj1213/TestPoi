package com.hxj.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by Administrator on 2017/8/18.
 */
public class Demo3 {

    public static void main(String[] args)
    {
        FileInputStream in = null;
        HSSFWorkbook workbook = null;


        try
        {
            in = new FileInputStream("sample.xls");
            POIFSFileSystem fs = new POIFSFileSystem(in);
            workbook = new HSSFWorkbook(fs);
        }
        catch (IOException e)
        {
            System.out.println(e.toString());
        }
        finally
        {
            try
            {
                in.close();
            }
            catch (IOException e)
            {
                System.out.println(e.toString());
            }
        }


        HSSFSheet sheet = workbook.getSheetAt(0);//取得第一张sheet
        int lastRowNum = sheet.getLastRowNum();
        System.out.println("---lastRowNum----"+lastRowNum);

for (int j = 0;j<4;j++){
    HSSFRow row = sheet.getRow(j);//第2行
    int lastCellNum = row.getLastCellNum();
    System.out.println("---lastCellNum----"+lastCellNum);
    for (int i = 0; i < 4; i++)
    {
        HSSFCell c = row.getCell((short) i);
        if (c == null)
        {
            System.out.println("第" + i + "列单元格不存在");
        }
        else
        {
            System.out.println("第" + i + "列单元格取得成功");

            if(j>0 && i==1){
                System.out.println("单元格的值：" + c.getNumericCellValue());//getStringCellValue()取得单元格的值
            }else{
                System.out.println("单元格的值：" + c.getStringCellValue());//getStringCellValue()取得单元格的值
            }
        }
    }
}

    }
}
