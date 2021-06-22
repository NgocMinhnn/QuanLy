/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package DAO;

import BUS.Tool;
import static DAO.KhachHangDAOTest.conn;
import DTO.CTHoaDonDTO;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
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
 * @author minh-pc
 */
public class ThongKeDAOTest {

    public static Connection conn;
    Workbook wb = null;

    public ThongKeDAOTest() {
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
            FileInputStream file = new FileInputStream("./TestFile/" + "THONGKEDAO.xlsx");
            wb = new XSSFWorkbook(file);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(BaoHanhDAOTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @After
    public void tearDown() {
        Tool.writeEx(wb, "THONGKEDAO.xlsx");
    }

  
//    @Test
//    public void testGetThuNhapNgay() {
//        System.out.println("getThuNhapNgay");
//        System.out.println(Tool.getCellNumber(wb, 0, 1, 10));
//         System.out.println(Tool.getCellStringValue(wb, 0, 5, 1));
//        System.out.println(Tool.getCellStringValue(wb, 0, 5, 2));
////        System.out.println(  ThongKeDAO.getThuNhapNgay(Tool.getCellStringValue(wb, 0, 5, 1), Tool.getCellStringValue(wb, 0, 5, 2)));
//        for(int i=5;i<Tool.getCellNumber(wb, 0, 1, 10)+5;i++   ){
//             String date1 = Tool.getCellStringValue(wb, 0, i, 1);
//             String date2 = Tool.getCellStringValue(wb, 0, i, 2);
//            
//            int result = ThongKeDAO.getThuNhapNgay(date1, date2);
//              System.out.println(result);  
//            if(result!=0){
//              
//                Tool.setValue(wb, 0, i, 5,result+" ");
//                  Tool.setValue(wb, 0, i, 3," !null ");
//            }
//            else{
//                Tool.setValue(wb, 0, i, 3, "null");
//                
//            }
//            
//        }
//       
//     
//    }

    
//    @Test
//    public void testGetThuNhapThang() {
//        System.out.println("getThuNhapThang");
//        String date1 = "";
//        String date2 = "";
//        String nam = "";
//        int expResult = 0;
//        int result = ThongKeDAO.getThuNhapThang(date1, date2, nam);
//        assertEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    }
  
//    @Test
//    public void testGetThuNhapNam() {
//        System.out.println("getThuNhapNam");
//        
//        System.out.println(Tool.getCellStringValue(wb, 2, 5, 2));
//       
//        for(int i=5;i<Tool.getCellNumber(wb, 2, 1, 10)+5;i++   ){
//        String nam = Tool.getCellStringValue(wb, 2, i, 2);
//         int result = ThongKeDAO.getThuNhapNam(nam);
//            System.out.println(result+"");
//           
//         
//         if(result!=0){
//                 Tool.setValue(wb, 2, i, 5, result+" ");
//                 Tool.setValue(wb, 2, i, 3," !null ");
//         }
//         else{  Tool.setValue(wb, 2, i, 3, "null");}
//        
//        }
//  
//    }

//   
//    @Test
//    public void testGetChiThang() {
//        System.out.println("getChiThang");
//        String date1 = "";
//        String date2 = "";
//        String nam = "";
//        int expResult = 0;
//        int result = ThongKeDAO.getChiThang(date1, date2, nam);
//        assertEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    }
//
//  
    @Test
    public void testGetChiNam() {
        System.out.println("getChiNam");
        
        for(int i=5;i<Tool.getCellNumber(wb, 0, 3, 10)+5;i++   ){
         String nam = Tool.getCellStringValue(wb, 3, i, 2);
         int result = ThongKeDAO.getChiNam(nam);
         
        if(result!=0){
             Tool.setValue(wb, 3, i, 5, result+" ");
                Tool.setValue(wb, 3, i, 3," !null ");
        }
        else{Tool.setValue(wb, 3, i, 3, "null");}
        }
       
     
     
       
        
     
    }


// 
//    @Test
//    public void testGetCTHOADON() {
//        System.out.println("getCTHOADON");
//        LocalDate date1 = null;
//        LocalDate date2 = null;
//        String sp = "";
//        String nv = "";
//        String KH = "";
//        ArrayList<CTHoaDonDTO> expResult = null;
//        ArrayList<CTHoaDonDTO> result = ThongKeDAO.getCTHOADON(date1, date2, sp, nv, KH);
//        assertEquals(expResult, result);
//        // TODO review the generated test code and remove the default call to fail.
//        fail("The test case is a prototype.");
//    }

}
