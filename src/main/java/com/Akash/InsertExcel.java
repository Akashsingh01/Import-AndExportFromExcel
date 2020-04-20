package com.Akash; 
import java.io.*;

    
    import org.apache.poi.ss.usermodel.*;
    import org.apache.poi.xssf.usermodel.XSSFSheet;
    import org.apache.poi.xssf.usermodel.XSSFWorkbook;
    

    import java.util.*;
    import java.sql.*; 

    public class InsertExcel {  
    	 public static   final String DB_DRIVER_CLASS="com.mysql.cj.jdbc.Driver";
    	 public static final String DB_USERNAME="root";
    	 public static 	final String DB_PASSWORD="root";
     	 public static final String DB_URL ="jdbc:mysql://localhost:3306/CrimsonLogic";

        public static final String INSERT_RECORDS = "INSERT INTO employees(id, name, department, project, domain, remarks) VALUES(?,?,?,?,?,?)";
        private static String GET_COUNT = "SELECT COUNT(*) FROM employees";

        public static void main(String[] args) throws Exception{
            InsertExcel obj = new InsertExcel();
            obj.insertRecords("C:/Users/engga/eclipse-workspace/Assignment1/UsersManager/Records.xlsx");



        }
              public void insertRecords(String filePath){

                /* Create Connection objects */
            Connection con = null;
            PreparedStatement prepStmt = null;
            java.sql.Statement stmt = null;
            int count = 0;
            ArrayList<String> mylist = new ArrayList<String>();
           

            try{
                con = DriverManager.getConnection(DB_URL, DB_USERNAME, DB_PASSWORD);
              //  System.out.println("Connection :: ["+con+"]");
                prepStmt = con.prepareStatement(INSERT_RECORDS);
                stmt = con.createStatement();
                ResultSet result = stmt.executeQuery(GET_COUNT);
                while(result.next()) {

                    int val = result.getInt(1);
                    System.out.println(val);
                    count = val+1;

                }


                //prepStmt.setInt(1,count);

                /* load excel objects and loop through the worksheet data */
                FileInputStream fis = new FileInputStream(new File(filePath));
                //System.out.println("FileInputStream Object created..! ");
                 /* Load workbook */
                XSSFWorkbook workbook = new XSSFWorkbook (fis);
             //   System.out.println("XSSFWorkbook Object created..! ");
                /* Load worksheet */
                XSSFSheet sheet = workbook.getSheetAt(0);
              //  System.out.println("XSSFSheet Object created..! ");
                   // loop through and insert data
                Iterator ite = sheet.rowIterator();
               // System.out.println("Row Iterator invoked..! ");

                   while(ite.hasNext()) {
                            Row row = (Row) ite.next(); 
                           // System.out.println("Row value fetched..! ");
                            Iterator<Cell> cellIterator = row.cellIterator();
                          //  System.out.println("Cell Iterator invoked..! ");
                            int index=1;
                                    while(cellIterator.hasNext()) {

                                            Cell cell = cellIterator.next();
                                           // System.out.println("getting cell value..! ");

                                           switch(cell.getCellType()) { 
                                       case STRING: //handle string columns
                                                    prepStmt.setString(index, cell.getStringCellValue());                                                                                     
                                                    break;
//                                            case Cell.CELL_TYPE_NUMERIC: //handle double data
//                                                int i = (int)cell.getNumericCellValue();
//                                                prepStmt.setInt(index, (int) cell.getNumericCellValue());
//                                                    break;
                                            }
                                            index++;



                                    }
                    //execute the statement before reading the next row
                    prepStmt.executeUpdate();
                    
                   /* Close input stream */
                   fis.close();
                   }
                   /* Close prepared statement */
                   prepStmt.close();

                   /* Close connection */
                   con.close();

            }catch(Exception e){
                e.printStackTrace();            
            }

            }
    }