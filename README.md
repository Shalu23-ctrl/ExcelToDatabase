# ExcelToDatabase
Excel spreadsheet to database
import example.Account;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.*;
import java.util.Iterator;

public class Excel_Db {
    XSSFRow row;
    Account acc = new Account();

    public static void main(String[] args) throws IOException {
        String fileName = "/home/saiteja/Documents/Account_Info.xlsx";
        Excel_Db Acc2 = new Excel_Db();
        Acc2.readFile(fileName);
    }

    public void readFile(String fileName) throws FileNotFoundException, IOException {
        FileInputStream fis;
        try {
            System.out.println("------READING THE SPREADSHEET-------");
            fis = new FileInputStream(fileName);
            XSSFWorkbook workbookRead = new XSSFWorkbook(fis);
            XSSFSheet spreadsheetRead = workbookRead.getSheetAt(0);
            Iterator<Row> rowIterator = spreadsheetRead.iterator();
            //Row headerRow = rowIterator.next(); // to remove header part from excel sheet
            int count = 0;
            while (rowIterator.hasNext()) {
                row = (XSSFRow) rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getColumnIndex()) {
                        case 0:
                            if (cell.getCellType().equals(CellType.NUMERIC)) {
                                System.out.print(
                                        cell.getNumericCellValue() + " \t\t");
                            } else if (cell.getCellType().equals(CellType.STRING)) {
                                System.out.print(
                                        cell.getStringCellValue() + " \t\t");
                            }
                            break;
                        case 1:
                            System.out.print(
                                    cell.getStringCellValue() + " \t\t");
                            break;
                        case 2:
                            if (cell.getCellType().equals(CellType.NUMERIC)) {
                                System.out.print(
                                        cell.getNumericCellValue() + " \t\t");
                            } else if (cell.getCellType().equals(CellType.STRING)) {
                                System.out.print(
                                        cell.getStringCellValue() + " \t\t");
                            }
                            break;
                        case 3:
                            if (cell.getCellType().equals(CellType.NUMERIC))
                                System.out.print(
                                        cell.getNumericCellValue() + " \t\t");
                            break;

                        case 4:
                            System.out.print(
                                    cell.getStringCellValue() + " \t\t");
                            break;
                    }
                    System.out.println();
                }
                if (count != 0) {
                    String id = row.getCell(0).getStringCellValue();

                     id=id.replace("_","");
                    acc.id = Integer.parseInt(id);
                    System.out.println(acc.id);
                    //}
                    acc.name = row.getCell(1).getStringCellValue();
                    if (row.getCell(2).getCellType().equals(CellType.NUMERIC)) {
                        acc.balance = Double.parseDouble(String.valueOf(row.getCell(2).getNumericCellValue()));
                    } else {
                        acc.balance = Double.parseDouble(row.getCell(2).getStringCellValue());
                    }
                    acc.Actionperformed = row.getCell(3).getStringCellValue();
                    if (acc.Actionperformed.equalsIgnoreCase("DELETE")) {
                        DeleteRowInDB(acc.id, acc.Actionperformed);
                        System.out.println("Values Deleted Succesfully");
                    } else if (acc.Actionperformed.equalsIgnoreCase("INSERT")) {
                        InsertRowInDB(acc.id, acc.name, acc.balance, acc.Actionperformed);
                        System.out.println("Values Inserted Successfully");
                    }
                    else if(acc.Actionperformed.equalsIgnoreCase("UPDATE")){
                        UpdateRowInDB(acc.id, acc.balance, acc.Actionperformed);
                        System.out.println("Values Updated Sucessfully");
                    }
                }
                ++count;
            }
            System.out.println();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public void InsertRowInDB(int id, String name, Double balance, String ActionPerformed) {
        try {
            Class.forName("org.postgresql.Driver");
            Connection c;
            c = DriverManager
                    .getConnection("jdbc:postgresql://localhost:5432/account",
                            "postgres", "whizzard@123");
            System.out.println("Opened database successfully");
            PreparedStatement ps = null;
            Statement stmt = c.createStatement();
            String sql = "INSERT INTO INFO (ID,NAME,BALANCE,ACTIONPERFORMED) VALUES(?,?,?,?)";
            ps = c.prepareStatement(sql);
            ps.setInt(1, id);
            ps.setString(2, name);
            ps.setDouble(3, balance);
            ps.setString(4, ActionPerformed);
            ps.executeUpdate();
            c.close();
        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }
    }

    public void UpdateRowInDB(int id,Double balance,String actionPerformed) {
        try {
            Class.forName("org.postgresql.Driver");
            Connection c;
            c = DriverManager
                    .getConnection("jdbc:postgresql://localhost:5432/account",
                            "postgres", "whizzard@123");
            System.out.println("Opened database successfully");
            PreparedStatement ps = null;
            String sql1 = "UPDATE INFO SET BALANCE=?,ACTIONPERFORMED=? WHERE ID=?";
            ps = c.prepareStatement(sql1);
            ps.setDouble(1, balance);
            ps.setString(2,actionPerformed);
            ps.setInt(3, id);
            ps.executeUpdate();   // Perform second update
            ps.close();
            c.close();
        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();
        }
    }

    public void DeleteRowInDB(int id, String status) {
        String UPDATE_STATUS_QUERY = "UPDATE INFO SET ACTIONPERFORMED=? WHERE ID=?";
        try {
            Class.forName("org.postgresql.Driver");
            Connection c;
            c = DriverManager
                    .getConnection("jdbc:postgresql://localhost:5432/account",
                            "postgres", "whizzard@123");
            System.out.println("Opened database successfully");
            PreparedStatement ps = null;
            ps = c.prepareStatement(UPDATE_STATUS_QUERY);
            ps.setString(1, status);
            ps.setInt(2, id);
            ps.executeUpdate();
            ps.close();
            c.close();
        } catch (ClassNotFoundException | SQLException e) {
            e.printStackTrace();

        }
    }

}
