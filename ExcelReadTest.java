import java.io.*;
import java.sql.*;
 
public class ExcelReadTest{
    public static void main(String[] args){    
        Connection connection = null;
        try{
            Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
            Connection con = DriverManager.getConnection( "jdbc:odbc:exceltest" );
            Statement st = con.createStatement();
            ResultSet rs = st.executeQuery( "Select * from [Sheet1$]" );
                 
            ResultSetMetaData rsmd = rs.getMetaData();
            int numberOfColumns = rsmd.getColumnCount();
                 
            while (rs.next()) {
                for (int i = 1; i <= numberOfColumns; i++) {
                    if (i > 1) 
                        System.out.print(", ");
                    String columnValue = rs.getString(i);
                    System.out.print(columnValue);
                }
                System.out.println(""); 
            }
                         
            st.close();
            con.close();
        } catch(Exception ex) {
            System.err.print("Exception: ");
            System.err.println(ex.getMessage());
        }
    }
}                