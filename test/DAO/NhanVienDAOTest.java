/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package DAO;

import BUS.Tool;

import DTO.NhanVienDTO;
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
public class NhanVienDAOTest {

    public static Connection conn;
    Workbook wb = null;

    public NhanVienDAOTest() {
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
            FileInputStream file = new FileInputStream("./TestFile/" + "NHANVIENDAO.xlsx");
            wb = new XSSFWorkbook(file);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(BaoHanhDAOTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @After
    public void tearDown() {
        Tool.writeEx(wb, "NHANVIENDAO.xlsx");
    }

    /**
     * Test of insertNhanVien method, of class NhanVienDAO.
     */
//    @Test
//    public void testInsertNhanVien() {
//        System.out.println("insertNhanVien");
//        ArrayList<NhanVienDTO> nv= new ArrayList();
//        
//        for(int i=5;i<Tool.getCellNumber(wb,0, 1, 10)+5;i++){
//            NhanVienDTO nvdto= new NhanVienDTO();
//            nvdto.setMaNhanVien(Tool.getCellStringValue(wb, 0, i, 1));
//            nvdto.setTenNhanVien(Tool.getCellStringValue(wb, 0, i, 2));
//            nvdto.setHoNhanVien(Tool.getCellStringValue(wb, 0, i, 3));
//            nvdto.setNgaySinh(Date.valueOf(Tool.getCellDate(wb, 0, i, 4)));
//            nvdto.setSdt(Tool.getCellStringValue(wb, 0, i, 5));
//            nvdto.setGioiTinh(Tool.getCellStringValue(wb, 0, i, 6));
//            nvdto.setMaQuyen(Tool.getCellStringValue(wb, 0, i, 7));
//            nvdto.setLuong((int)Tool.getCellNumber(wb, 0, i, 8));
//            nvdto.setTrangThai((int)Tool.getCellNumber(wb, 0, i, 9));
//                    
//          nv.add(nvdto);
//            
//        }
//        
//        
//        for(NhanVienDTO nvdt: nv){
//       
//       
//        boolean result = NhanVienDAO.insertNhanVien(nvdt);
//        
//        for(int i=5;i<Tool.getCellNumber(wb,0, 1, 10)+5;i++){
//             boolean expResult = Tool.getCellBooleanValue(wb, 0, i, 10);
//             if (result==expResult){
//                 Tool.setValue(wb, 0, i, 11, "TRUE");
//             }
//             
//             else{
//              fail("The test case is a prototype.");
//              Tool.setValue(wb, 0, i, 11, "FALSE");
//             }
//             
//        }
//       
//     
//       
//        }
//        
//    }

    /**
     * Test of getNhanVien method, of class NhanVienDAO.
     */
//    @Test
//    public void testGetNhanVien() {
//        System.out.println("getNhanVien");
//        
//        for(int i=5; i< Tool.getCellNumber(wb, 1, 1, 10)+5;i++){
//            ArrayList<NhanVienDTO> result = NhanVienDAO.getNhanVien();
//           String expResult= Tool.getCellStringValue(wb, 1, i, 3);
//           
//           if(result != null){
//               Tool.setValue(wb, 1, i, 4, "!null");
//           }
//           else {Tool.setValue(wb, 1, i, 4, "null");
//           fail("The test case is a prototype.");}
//           
//        }
    //  }

    /**
     * Test of updateNhanVien method, of class NhanVienDAO.
     */
//    @Test
//    public void testUpdateNhanVien() {
//        System.out.println("test updta nhan vien ");
//        ArrayList<NhanVienDTO> nv= new ArrayList();
//        
//        for(int i=5;i<Tool.getCellNumber(wb,2, 1, 10)+5;i++){
//            NhanVienDTO nvdto= new NhanVienDTO();
//            nvdto.setMaNhanVien(Tool.getCellStringValue(wb, 2, i, 1));
//            nvdto.setTenNhanVien(Tool.getCellStringValue(wb, 2, i, 2));
//            nvdto.setHoNhanVien(Tool.getCellStringValue(wb, 2, i, 3));
//            nvdto.setNgaySinh(Date.valueOf(Tool.getCellDate(wb, 2, i, 4)));
//            nvdto.setSdt(Tool.getCellStringValue(wb, 2, i, 5));
//            nvdto.setGioiTinh(Tool.getCellStringValue(wb, 2, i, 6));
//            nvdto.setMaQuyen(Tool.getCellStringValue(wb, 2, i, 7));
//            nvdto.setLuong((int)Tool.getCellNumber(wb, 2, i, 8));
//            nvdto.setTrangThai((int)Tool.getCellNumber(wb, 2, i, 9));
//                    
//          nv.add(nvdto);
//            
//        }
//        
//        for(NhanVienDTO nvdt: nv){
//            System.out.println(nvdt.getMaNhanVien());
//       
//         boolean result = NhanVienDAO.updateNhanVien(nvdt);
//        
//        for(int i=5;i<Tool.getCellNumber(wb,2, 1, 10)+5;i++){
//             boolean expResult = Tool.getCellBooleanValue(wb, 2, i, 10);
//             if (result==expResult){
//                 Tool.setValue(wb, 2, i, 11, "TRUE");
//             }
//             
//             else{
//              fail("The test case is a prototype.");
//              Tool.setValue(wb, 2, i, 11, "FALSE");
//             }
//             
//     }
// }
//}
      @Test
    public void testDeleteNhanVien() {
        System.out.println("DeleteNhanVien");
        ArrayList<NhanVienDTO> kh = new ArrayList();
        
        for (int i = 5; i < Tool.getCellNumber(wb, 3, 1, 10) + 5; i++) {
            NhanVienDTO khdto = new NhanVienDTO();
            khdto.setMaNhanVien(Tool.getCellStringValue(wb, 3, i, 1));
            kh.add(khdto);
        }
        for (NhanVienDTO khd : kh) {
            System.out.println(khd.getMaNhanVien());
            boolean result = NhanVienDAO.DeleteNhanVien(khd);
            for (int i = 5; i < Tool.getCellNumber(wb, 3, 1, 10) + 5; i++) {
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