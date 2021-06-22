/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package DAO;

import BUS.Tool;
import static DAO.KhachHangDAOTest.conn;
import DTO.TaiKhoanDTO;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author minh-pc
 */
public class TaiKhoanDAOTest {

    public static Connection conn;
    Workbook wb = null;

    public TaiKhoanDAOTest() throws SQLException {
        try {
            conn = MySQLConnUtils.getMySQLConnection();
        } catch (ClassNotFoundException ex) {
            Logger.getLogger(BaoHanhDAOTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @BeforeClass
    public static void setUpClass() throws SQLException {
         conn.close();
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() throws IOException {
         try {
            FileInputStream file = new FileInputStream("./TestFile/" + "TAIKHOANDAO.xlsx");
            wb = new XSSFWorkbook(file);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(BaoHanhDAOTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @After
    public void tearDown() {
          Tool.writeEx(wb, "TAIKHOANDAO.xlsx");
    }

    /**
     * Test of insertTaiKhoan method, of class TaiKhoanDAO.
     */
    @Test
    public void testInsertTaiKhoan() {
        System.out.println("insertTaiKhoan");
        
        
        TaiKhoanDTO sp = null;
        boolean expResult = false;
        boolean result = TaiKhoanDAO.insertTaiKhoan(sp);
        assertEquals(expResult, result);
        fail("The test case is a prototype.");
    }

    /**
     * Test of getTaiKhoan method, of class TaiKhoanDAO.
     */
    @Test
    public void testGetTaiKhoan() {
        System.out.println("getTaiKhoan");
        ArrayList<TaiKhoanDTO> expResult = null;
        ArrayList<TaiKhoanDTO> result = TaiKhoanDAO.getTaiKhoan();
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of updateTaiKhoan method, of class TaiKhoanDAO.
     */
    @Test
    public void testUpdateTaiKhoan() {
        System.out.println("updateTaiKhoan");
        TaiKhoanDTO sp = null;
        boolean expResult = false;
        boolean result = TaiKhoanDAO.updateTaiKhoan(sp);
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of DeleteTaiKhoan method, of class TaiKhoanDAO.
     */
    @Test
    public void testDeleteTaiKhoan() {
        System.out.println("DeleteTaiKhoan");
        TaiKhoanDTO sp = null;
        boolean expResult = false;
        boolean result = TaiKhoanDAO.DeleteTaiKhoan(sp);
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

    /**
     * Test of main method, of class TaiKhoanDAO.
     */
    @Test
    public void testMain() {
        System.out.println("main");
        String[] args = null;
        TaiKhoanDAO.main(args);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

}
