package com.Akash;

import java.io.*;
import java.sql.*;
import java.util.*;
import java.util.Date;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
 
public class ReadWriteExcel {

    public static void main(String ar[]) {
        ReadWriteExcel rw = new ReadWriteExcel();
        rw.readDataFromExcel();

    }
    Object[][] data = null;

    public File getFile() throws FileNotFoundException {
        File here = new File("C:\\Users\\engga\\eclipse-workspace\\Assignment1\\UsersManager\\Book1.xlsx");
        return new File(here.getAbsolutePath());

    }

    public Object[][] readDataFromExcel() {
        final DataFormatter df = new DataFormatter();
        try {

            FileInputStream file = new FileInputStream(getFile());
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();

            int rownum = 0;
            int colnum = 0;
            Row r=rowIterator.next();

            int rowcount=sheet.getLastRowNum();
            int colcount=r.getPhysicalNumberOfCells();
            data = new Object[rowcount][colcount];
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                colnum = 0;
                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                    data[rownum][colnum] =  df.formatCellValue(cell);
                    System.out.print(df.formatCellValue(cell));
                    colnum++;
                    System.out.print("");
                }
                rownum++;
                System.out.println(" ");
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        return data;
    }
}