package com.Akash;

import java.io.*;
import java.sql.*;
 
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
 

public class StoreExcel {
	private static final String DB_DRIVER_CLASS="com.mysql.cj.jdbc.Driver";
	private static final String DB_USERNAME="root";
	private static final String DB_PASSWORD="root";
	private static final String DB_URL ="jdbc:mysql://localhost:3306/CrimsonLogic";
	private static Connection connection = null;
 
    public static void main(String[] args) {
    	
    	
    	
        new StoreExcel().export();
    }
     
    public void export() {
//        String jdbcURL = "jdbc:mysql://localhost:3306/CrimsonLogic";
//        String username = "root";
//        String password = "root";
 
        String excelFilePath = "C:\\Users\\engga\\Desktop\\export.xlsx";
 
        try { connection = DriverManager.getConnection(DB_URL, DB_USERNAME, DB_PASSWORD);
        		
            String sql = "SELECT  * FROM employees ORDER BY id DESC limit 1";
 
            Statement statement = connection.createStatement();
 
            ResultSet result = statement.executeQuery(sql);
 
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("DataInfo");
 
            writeHeaderLine(sheet);
 
            writeDataLines(result, workbook, sheet);
 
            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
            System.out.println("Data Succefully export");
            workbook.close();
 
            statement.close();
 
        }
    catch (SQLException e) {
            System.out.println("Datababse error:");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("File IO error:");
            e.printStackTrace();
        }
    }
 
    private void writeHeaderLine(XSSFSheet sheet) {
 
        Row headerRow = sheet.createRow(0);
 
        Cell headerCell = headerRow.createCell(0);
        headerCell.setCellValue("Id");
 
        headerCell = headerRow.createCell(1);
        headerCell.setCellValue("Name");
 
        headerCell = headerRow.createCell(2);
        headerCell.setCellValue("Department");
 
        headerCell = headerRow.createCell(3);
        headerCell.setCellValue("Project");
 
        headerCell = headerRow.createCell(4);
        headerCell.setCellValue("Domain");
        headerCell = headerRow.createCell(5);
        headerCell.setCellValue("remarks");
    }
 
    private void writeDataLines(ResultSet result, XSSFWorkbook workbook,
            XSSFSheet sheet) throws SQLException {
        int rowCount = 1;
 
        while (result.next()) {
        	int id =result.getInt("Id");
            String Name = result.getString("Name");
            String department = result.getString("Department");
            String project = result.getString("Project");
            String domain = result.getString("Domain");
           
            String remarks = result.getString("remarks");
 
            Row row = sheet.createRow(rowCount++);
 
            int columnCount = 0;
            Cell cell = row.createCell(columnCount++);
            cell.setCellValue(id);
 
            cell = row.createCell(columnCount++);
            cell.setCellValue( Name);
 
            cell = row.createCell(columnCount++);
            cell.setCellValue( department);
 
            cell = row.createCell(columnCount++);
            cell.setCellValue( project);
            cell = row.createCell(columnCount++);
            cell.setCellValue( domain);
           
            cell = row.createCell(columnCount++);
            cell.setCellValue( remarks);
           
           
 
          
 
         
 
        }
    }
 
}