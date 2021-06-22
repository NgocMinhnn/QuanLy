/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package DAO;

import BUS.BaoHanhBUS;
import BUS.CTHDBUS;
import BUS.HoaDonBUS;
import BUS.Tool;
import DTO.BaoHanhDTO;
import DTO.CTHoaDonDTO;
import DTO.HoaDonDTO;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.Date;
import java.sql.SQLException;
import java.time.LocalDate;
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
 * @author Vind
 */
public class BaoHanhDAOTest {

    public static Connection conn;
    Workbook wb = null;

    public BaoHanhDAOTest() {
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
            FileInputStream file = new FileInputStream("./TestFile/" + "BAOHANHDAO.xlsx");
            wb = new XSSFWorkbook(file);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(BaoHanhDAOTest.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    @After
    public void tearDown() {
        Tool.writeEx(wb, "BAOHANHDAO.xlsx");
    }

    /**
     * Test of insertBaoHanh method, of class BaoHanhDAO.
     */
    @Test
    public void testInsertBaoHanh() {
        System.out.println("insertBaoHanh");
        String mahd = Tool.getCellStringValue(wb, 0, 5, 1);
        System.out.println(Tool.getCellStringValue(wb, 0, 5, 1));
        HoaDonBUS hdb= new HoaDonBUS();
        CTHDBUS cthdb= new CTHDBUS();
        BaoHanhDTO sp = null;
        
        HoaDonDTO hd ;
        hd=hdb.getHD(mahd);
        System.out.println(hd.getMaHoaDon());
        CTHoaDonDTO ct;
        ct=cthdb.cthd(hd.getMaHoaDon());
        System.out.println(ct.getMaSanPham());
        sp=new BaoHanhDTO(hd.getMaHoaDon(),ct.getMaSanPham(),Date.valueOf("2021-05-09"),12);
        
        
        boolean expResult = Tool.getCellBooleanValue(wb, 0, 5, 3);
        System.out.println(expResult);
        boolean result = BaoHanhDAO.insertBaoHanh(sp);
        System.out.println(result);
        
        assertEquals(expResult, result);
        Tool.setBooleanValue(wb, 0, 5, 4, result);
        // TODO review the generated test code and remove the default call to fail.
     
    }
    /**
     * Test of getBaoHanh method, of class BaoHanhDAO.
     */
    @Test
    public void testGetBaoHanh() {
        System.out.println("getBaoHanh");
        String exResult= Tool.getCellStringValue(wb, 1,5,3);
        System.out.println("EO: "+Tool.getCellStringValue(wb, 1,5,3));
        System.out.println( Tool.getCellNumber(wb, 1, 1, 10));
        for(int i=5;i<Tool.getCellNumber(wb, 1, 1, 10)+5;i++){
            ArrayList<BaoHanhDTO> result = BaoHanhDAO.getBaoHanh();
        if ( result !=null){
            
                assertEquals(true, true);
                System.out.println("AO: !null" );
                Tool.setValue(wb, 1, i, 4, "!null");
        }
        else{
            fail("The test case is a prototype.");
            System.out.println("AO: null ");
            Tool.setValue(wb, 1, i, 4, "null");
        }
        
        }
        
        
       
    }

    @Test
    public void testUpdateBaoHanh() {
        System.out.println("updateBaoHanh");
        ArrayList<BaoHanhDTO> bh = new ArrayList();
        BaoHanhDTO sp = null;
        System.out.println(Tool.getCellNumber(wb, 2, 1, 10));
        System.out.println(Tool.getCellDate(wb, 2, 5, 1));

        for (int i = 5; i < Tool.getCellNumber(wb, 2, 1, 10) + 5; i++) {

            BaoHanhDTO bhdto = new BaoHanhDTO(Tool.getCellStringValue(wb, 2, i, 3),
                    Tool.getCellStringValue(wb, 2, i, 4),
                    (Date.valueOf(Tool.getCellDate(wb, 2, 5, 1))),
                    (int) Tool.getCellNumber(wb, 2, i, 2));
            bh.add(bhdto);
        }
        for (BaoHanhDTO bhdto : bh) {
            for (int i = 5; i < Tool.getCellNumber(wb, 2, 1, 10) + 5; i++) {
                boolean result = BaoHanhDAO.updateBaoHanh(bhdto);
                boolean expResult = Tool.getCellBooleanValue(wb, 2, i, 5);
                assertEquals(expResult, result);
                Tool.setBooleanValue(wb, 2, i, 6, result);

            }

        }

        
    }

}
