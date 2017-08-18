package com.hxj.util;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;

/**
 * Created by Administrator on 2017/8/18.
 */
public class Demo8 {

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


            HSSFSheet sheet = workbook.getSheetAt(0);
            HSSFRow row = sheet.getRow(1);


            System.out.println("First:" + row.getFirstCellNum());
            System.out.println("Last:" + row.getLastCellNum());
            System.out.println("Total:" + row.getPhysicalNumberOfCells() + "\n");


            row = sheet.getRow(2);

            System.out.println("First:" + row.getFirstCellNum());
            System.out.println("Last:" + row.getLastCellNum());
            System.out.println("Total:" + row.getPhysicalNumberOfCells() + "\n");

        }
}
