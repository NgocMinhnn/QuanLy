/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package DAO;

import BUS.Tool;
import static DAO.KhachHangDAOTest.conn;
import DTO.SanPhamDTO;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Date;
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
public class SanPhamDAOTest {

    public static Connection conn;
    Workbook wb = null;

    public SanPhamDAOTest() {
    }

    @BeforeClass
    public static void setUpClass() throws SQLException {
        try {
            conn = MySQLConnUtils.getMySQLConnection();
        } catch (ClassNotFoundException ex) {
            Logger.getLogger(BaoHanhDAOTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @AfterClass
    public static void tearDownClass() throws SQLException {
        conn.close();
    }

    @Before

    public void setUp() throws IOException {
        try {
            FileInputStream file = new FileInputStream("./TestFile/" + "SANPHAMDAO.xlsx");
            wb = new XSSFWorkbook(file);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(BaoHanhDAOTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @After
    public void tearDown() {
        Tool.writeEx(wb, "SANPHAMDAO.xlsx");
    }

    /**
     * Test of insertSanPham method, of class SanPhamDAO.
     */
    @Test
    public void testInsertSanPham() {
        System.out.println("insertSanPham");
        ArrayList<SanPhamDTO> sp = new ArrayList();
        System.out.println(Tool.getCellStringValue(wb, 0, 5, 1));
      
        for (int i = 5; i < Tool.getCellNumber(wb, 0, 1, 10) + 5; i++) {

            SanPhamDTO spdto = new SanPhamDTO();
            spdto.setMaSanPham(Tool.getCellStringValue(wb, 0, i, 1));
            spdto.setTenSanPham(Tool.getCellStringValue(wb, 0, i, 2));
            spdto.setDongia((int) Tool.getCellNumber(wb, 0, i, 3));
            spdto.setNamSx(Date.valueOf(Tool.getCellDate(wb, 0, i, 4)));
            spdto.setSoLuong((int) Tool.getCellNumber(wb, 0, i, 5));
            spdto.setMota(Tool.getCellStringValue(wb, 0, i, 6));
            spdto.setMaLoaiSP(Tool.getCellStringValue(wb, 0, i, 7));
            spdto.setMaThuongHieu(Tool.getCellStringValue(wb, 0, i, 8));
            sp.add(spdto);

        }

        for (SanPhamDTO spd : sp) {
           
            boolean result = SanPhamDAO.insertSanPham(spd);
             for (int i = 5; i < Tool.getCellNumber(wb, 0, 1, 10) + 5; i++) {
                   boolean expResult = Tool.getCellBooleanValue(wb, 0, i, 9);
                   
                   if (result==expResult){
                       Tool.setValue(wb, 0, i, 10, "TRUE");
                   }
                   else{
                       Tool.setValue(wb, 0, i, 10, "FALSE");
                    fail("The test case is a prototype.");
                   }
             }
        }

    }

    /**
     * Test of getSanPham method, of class SanPhamDAO.
     */
    @Test
    public void testGetSanPham() {
        System.out.println("getSanPham");
        
          for (int i = 5; i < Tool.getCellNumber(wb, 1, 1, 10); i++) {
        String expResult = Tool.getCellStringValue(wb, 1, i, 3);
        ArrayList<SanPhamDTO> result = SanPhamDAO.getSanPham();
        if(result!=null){
                assertEquals(true, true);
                System.out.println("AO: !null");
                Tool.setValue(wb, 1, i, 4, "!null");
        }
        else{   fail("The test case is a prototype.");
                System.out.println("AO: null ");
                Tool.setValue(wb, 1, i, 4, "null");}
          }
     
    }

    /**
     * Test of updateSanPham method, of class SanPhamDAO.
     */
    @Test
    public void testUpdateSanPham() {
        System.out.println("updateSanPham");
         ArrayList<SanPhamDTO> sp = new ArrayList();
        System.out.println(Tool.getCellStringValue(wb, 0, 5, 1));
      
        for (int i = 5; i < Tool.getCellNumber(wb, 0, 1, 10) + 5; i++) {

            SanPhamDTO spdto = new SanPhamDTO();
            spdto.setMaSanPham(Tool.getCellStringValue(wb, 0, i, 1));
            spdto.setTenSanPham(Tool.getCellStringValue(wb, 0, i, 2));
            spdto.setDongia((int) Tool.getCellNumber(wb, 0, i, 3));
            spdto.setNamSx(Date.valueOf(Tool.getCellDate(wb, 0, i, 4)));
            spdto.setSoLuong((int) Tool.getCellNumber(wb, 0, i, 5));
            spdto.setMota(Tool.getCellStringValue(wb, 0, i, 6));
            spdto.setMaLoaiSP(Tool.getCellStringValue(wb, 0, i, 7));
            spdto.setMaThuongHieu(Tool.getCellStringValue(wb, 0, i, 8));
            sp.add(spdto);

        }
            for (SanPhamDTO spd : sp) {
           
           boolean result = SanPhamDAO.updateSanPham(spd);
             for (int i = 5; i < Tool.getCellNumber(wb, 0, 1, 10) + 5; i++) {
                   boolean expResult = Tool.getCellBooleanValue(wb, 0, i, 9);
                   
                   if (result==expResult){
                       Tool.setValue(wb, 0, i, 10, "TRUE");
                   }
                   else{
                       Tool.setValue(wb, 0, i, 10, "FALSE");
                    fail("The test case is a prototype.");
                   }
             }
        }
        
       
    }

    /**
     * Test of DeleteSanPham method, of class SanPhamDAO.
     */
    @Test
    public void testDeleteSanPham() {
        System.out.println("DeleteSanPham");
        SanPhamDTO sp = null;
        boolean expResult = false;
        boolean result = SanPhamDAO.DeleteSanPham(sp);
        assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        fail("The test case is a prototype.");
    }

}
