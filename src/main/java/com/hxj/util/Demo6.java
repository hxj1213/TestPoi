package com.hxj.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by Administrator on 2017/8/18.
 */
public class Demo6
{
    public static void main(String[] args){
        FileInputStream in = null;
        HSSFWorkbook workbook = null;


        try
        {
            in = new FileInputStream("sample1.xls");
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


        HSSFSheet sheet = workbook.getSheetAt(0);
        HSSFRow row = sheet.getRow(1);


        HSSFCell cell = row.getCell((short) 0);
        System.out.println("A:2=" + getType(cell)+"   "+cell.getBooleanCellValue());


        cell = row.getCell((short) 1);
        System.out.println("B:2=" + getType(cell)+"    "+cell.getDateCellValue()+"   "+HSSFDateUtil.isCellDateFormatted(cell));


        cell = row.getCell((short) 2);
        System.out.println("C:2=" + getType(cell));


        cell = row.getCell((short) 3);
        System.out.println("D:2=" + getType(cell));


        cell = row.getCell((short) 4);
        System.out.println("E:2=" + getType(cell));
    }


    public static String getType(HSSFCell cell)
    {
        int type = cell.getCellType();


        if (type == HSSFCell.CELL_TYPE_BLANK)
        {
            return "CELL_TYPE_BLANK";
        }
        else if (type == HSSFCell.CELL_TYPE_BOOLEAN)
        {
            return "CELL_TYPE_BOOLEAN";
        }
        else if (type == HSSFCell.CELL_TYPE_ERROR)
        {
            return "CELL_TYPE_ERROR";
        }
        else if (type == HSSFCell.CELL_TYPE_FORMULA)
        {
            return "CELL_TYPE_FORMULA";
        }
        else if (type == HSSFCell.CELL_TYPE_NUMERIC)
        {
            System.out.println("----------------"+HSSFDateUtil.isCellDateFormatted(cell));
            //检查日期类型
            if (HSSFDateUtil.isCellDateFormatted(cell))
            {
                System.out.println("***********");
                return "CELL_TYPE_DATE";
            }
            else
            {
                return "CELL_TYPE_NUMERIC";
            }
        }
        else if (type == HSSFCell.CELL_TYPE_STRING)
        {
            return "CELL_TYPE_STRING";
        }
        else
        {
            return "Not defined";
        }
    }

}
