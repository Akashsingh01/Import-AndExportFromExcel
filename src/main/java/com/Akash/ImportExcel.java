package com.Akash;

// Directly import data from excell
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.sql.ResultSet;
import java.sql.SQLException;

@SuppressWarnings("unused")
public class ImportExcel
{
   public static void main(String[] args) 
   {
       DBase db = new DBase();
       Connection conn = db.connect(
   "jdbc:mysql://localhost:3306/CrimsonLogic","root","root");
//       db.importData(conn,args[0]);
       db.importData(conn, "C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/fsf.xls");
   }

}

class DBase
{
   public DBase()
   {
   }

   public Connection connect(String db_connect_str, 
 String db_userid, String db_password)
   {
       Connection conn;
       try 
       {
           Class.forName(  
   "com.mysql.jdbc.Driver").newInstance();

           conn = DriverManager.getConnection(db_connect_str, 
   db_userid, db_password);

       }
       catch(Exception e)
       {
           e.printStackTrace();
           conn = null;
       }

       return conn;    
   }

   public void importData(Connection conn,String filename)
   {
       Statement stmt;
       String query;

       try
       {
           stmt = conn.createStatement(
   ResultSet.TYPE_SCROLL_SENSITIVE,
   ResultSet.CONCUR_UPDATABLE);

           query = "LOAD DATA  INFILE  'C:/ProgramData/MySQL/MySQL Server 8.0/Uploads/fsf.xls' INTO TABLE employees (id,name,department,project,domain,remarks);";

           stmt.executeUpdate(query);

       }
       catch(Exception e)
       {
           e.printStackTrace();
           stmt = null;
       }
   }
};