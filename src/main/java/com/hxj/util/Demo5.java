package com.hxj.util;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

/**
 * Created by Administrator on 2017/8/18.
 */
public class Demo5
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
        System.out.println("A:2=" + getType(cell.getCellType()));


        cell = row.getCell((short) 1);
        System.out.println("B:2=" + getType(cell.getCellType()));


        cell = row.getCell((short) 2);
        System.out.println("C:2=" + getType(cell.getCellType()));


        cell = row.getCell((short) 3);
        System.out.println("D:2=" + getType(cell.getCellType()));


        cell = row.getCell((short) 4);
        System.out.println("E:2=" + getType(cell.getCellType()));
    }

    public static String getType(int type)
    {
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
            return "CELL_TYPE_NUMERIC";
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
