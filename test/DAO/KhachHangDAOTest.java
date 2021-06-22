/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package DAO;

import BUS.Tool;
import static DAO.BaoHanhDAOTest.conn;
import DTO.KhachHangDTO;
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
public class KhachHangDAOTest {
    public static Connection conn;
    Workbook wb = null;


    public KhachHangDAOTest() {
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
            FileInputStream file = new FileInputStream("./TestFile/" + "KHACHHANGDAO.xlsx");
            wb = new XSSFWorkbook(file);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(BaoHanhDAOTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @After
    public void tearDown() {
        Tool.writeEx(wb, "KHACHHANGDAO.xlsx");
    }

    @Test
    public void testInsertKhachHang() {
        System.out.println("insertKhachHang");
        KhachHangDTO sp = null;
       
        ArrayList<KhachHangDTO> kh= new ArrayList();
     
        System.out.println(Tool.getCellNumber(wb, 0, 1, 10));
        for (int i=5; i<Tool.getCellNumber(wb, 0, 1, 10)+5;i++ ){
            System.out.println(Tool.getCellStringValue(wb, 0, i, 1));
            System.out.println(Tool.getCellStringValue(wb, 0, i, 2));
            System.out.println(Tool.getCellStringValue(wb, 0, i, 3));
            System.out.println(Date.valueOf(Tool.getCellDate(wb, 0, i, 4)));
            System.out.println(Tool.getCellStringValue(wb, 0, i, 5));
            System.out.println(Tool.getCellStringValue(wb, 0, i, 6));
            System.out.println((int)Tool.getCellNumber(wb, 0, i, 7));

            KhachHangDTO khdto = new KhachHangDTO();
            khdto.setMaKhachHang(Tool.getCellStringValue(wb, 0, i, 1));
            khdto.setTenKhachHang(Tool.getCellStringValue(wb, 0, i, 2));
            khdto.setHoKhachHang(Tool.getCellStringValue(wb, 0, i, 3));
            khdto.setNgaySinh(Date.valueOf(Tool.getCellDate(wb, 0, i, 4)));
            khdto.setSdt(Tool.getCellStringValue(wb, 0, i, 5));
            khdto.setLoaiKhachHang(Tool.getCellStringValue(wb, 0, i, 6));
            khdto.setTichLuy((int)Tool.getCellNumber(wb, 0, i, 7));
            
            kh.add(khdto);
        }
       
      
        for(KhachHangDTO khd :kh){
        boolean result = KhachHangDAO.insertKhachHang(khd);
            System.out.println(result);
             for (int i=5; i<Tool.getCellNumber(wb, 0, 1, 10)+5;i++ ){
               boolean  exResult=Tool.getCellBooleanValue(wb, 0, i, 8);
                 assertEquals(exResult, result);
                 Tool.setBooleanValue(wb, 0, i, 9, result);                
             }          
        }            
    }
    @Test
    public void testGetKhachHang() {
        System.out.println("getKhachHang");

        for (int i = 5; i < Tool.getCellNumber(wb, 1, 1, 10); i++) {
            String expResult = Tool.getCellStringValue(wb, 1, i, 3);
            ArrayList<KhachHangDTO> result = KhachHangDAO.getKhachHang();

            if (result != null) {

                assertEquals(true, true);
                System.out.println("AO: !null");
                Tool.setValue(wb, 1, i, 4, "!null");
            } else {
                fail("The test case is a prototype.");
                System.out.println("AO: null ");
                Tool.setValue(wb, 1, i, 4, "null");

            }
        }

    }

    @Test
    public void testUpdateKhachHang() {
        System.out.println("updateKhachHang");
        
        ArrayList<KhachHangDTO> kh = new ArrayList();
    
        for (int i = 5; i < Tool.getCellNumber(wb, 2, 1, 10) + 5; i++) {
            KhachHangDTO khdto = new KhachHangDTO();
            khdto.setMaKhachHang(Tool.getCellStringValue(wb, 2, i, 1));
            khdto.setHoKhachHang(Tool.getCellStringValue(wb, 2, i, 2));
            khdto.setTenKhachHang(Tool.getCellStringValue(wb, 2, i, 3));
            khdto.setNgaySinh(Date.valueOf(Tool.getCellDate(wb, 2, i, 4)));
            khdto.setSdt(Tool.getCellStringValue(wb, 2, i, 5));
            khdto.setLoaiKhachHang(Tool.getCellStringValue(wb, 2, i, 6));
            khdto.setTichLuy((int) Tool.getCellNumber(wb, 2, i, 7));
            kh.add(khdto);

        }
       
        for (KhachHangDTO khdto : kh) {
          
            boolean result = KhachHangDAO.updateKhachHang(khdto);
            System.out.println(result);
            for (int i = 5; i < Tool.getCellNumber(wb, 2, 1, 10) + 5; i++) {
                boolean expResult = Tool.getCellBooleanValue(wb, 2, i, 8);
           
                if (result == expResult) {
                        Tool.setValue(wb, 2, i, 9, "TRUE");
                }
                else{ 
                    Tool.setValue(wb, 2, i, 9, "FALSE");
                    fail("The test case is a prototype.");}
            }

           
        }

    }
    @Test
    public void testDeleteKhachHang() {
        System.out.println("DeleteKhachHang");
        ArrayList<KhachHangDTO> kh = new ArrayList();
        System.out.println(Tool.getCellNumber(wb, 3, 1, 10) );
        for (int i = 5; i < Tool.getCellNumber(wb, 3, 1, 10) + 5; i++) {
            KhachHangDTO khdto = new KhachHangDTO();
            khdto.setMaKhachHang(Tool.getCellStringValue(wb, 3, i, 1));
            kh.add(khdto);
        }
        for (KhachHangDTO khd : kh) {
            System.out.println(khd.getMaKhachHang());
            boolean result = KhachHangDAO.DeleteKhachHang(khd);
            for (int i = 5; i < Tool.getCellNumber(wb, 2, 1, 10) + 5; i++) {
                boolean expResult = Tool.getCellBooleanValue(wb, 3, i, 3);
               if (result==expResult){
                   Tool.setValue(wb, 3, i, 4, "TRUE");
               }
               else{
                    Tool.setValue(wb, 3, i, 4, "FALSE");
                    fail("The test case is a prototype.");
               }
            }

           
        }

    }

}
