/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package BUS;

import DTO.BaoHanhDTO;
import java.time.LocalDate;
import java.util.ArrayList;
import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;
import BUS.Tool;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Vind
 */
public class BaoHanhBUSTest {

    String filePath = "BHBUS.xlsx";

    public BaoHanhBUSTest() {
    }

    @BeforeClass
    public static void setUpClass() {
    }

    @AfterClass
    public static void tearDownClass() {
    }

    @Before
    public void setUp() {
    }

    @After
    public void tearDown() {
    }

    /**
     * Test of getData method, of class BaoHanhBUS.
     */
    @Test
    public void testGetData() {
        System.out.println("getData");
        BaoHanhBUS instance = new BaoHanhBUS();
        try {
            instance.getData();
            assertFalse(instance.list.equals(null));
        } catch (Exception e) {
            fail("Cant Get Data from DB");
        }
    }

    /**
     * Test of getds method, of class BaoHanhBUS.
     */
    @Test
    public void testGetds() {
        System.out.println("getds");
        BaoHanhBUS instance = new BaoHanhBUS();
        Workbook wb = null;
        try{
            wb = new XSSFWorkbook("./TestFile/"+filePath);
        }catch(Exception e){
            e.printStackTrace();
        }
        for(int i=5;i<Tool.getCellNumber(wb, 1, 1, 5)+5;i++){
            try {
                ArrayList<BaoHanhDTO> result = instance.getds();
                if(result!=null){
                    assertEquals(true, true);
                    Tool.setValue(wb, 1, i, 3, "!null");
                    System.out.println(Tool.getCellNumber(wb, 1, 1, 5));
                }
            } catch (Exception e) {
                    fail("Cant get ds");
                    Tool.setValue(wb, 1, i, 3, "null");
            }
        }
        
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }

    /**
     * 
     * Test of search method, of class BaoHanhBUS.
     */
//    @Test
//    public void testSearch() {
//        System.out.println("search");
//        BaoHanhBUS instance = new BaoHanhBUS();
//        Workbook wb = null;
//        try{
//            wb = new XSSFWorkbook("./TestFile/"+filePath);
//        }catch(Exception e){
//            e.printStackTrace();
//        }
//        for (int i = 5; i <Tool.getCellNumber(wb, 0, 1, 8); i++) {
//            String value = Tool.getCellStringValue(wb,0, i, 1);
//            String Type = Tool.getCellStringValue(wb,0, i, 2);
//            LocalDate ngay_1 = Tool.getCellDate(wb,0, i, 3);
//            LocalDate ngay_2 = Tool.getCellDate(wb,0, i, 4);
//            String expResult = Tool.getCellStringValue(wb,0, i, 5);
//            ArrayList<BaoHanhDTO> result = instance.search(value, Type, null, null);
//            if (result != null) {
//                assertEquals(result.get(0).getMaHoaDon(), expResult);
//                Tool.setValue(wb, 0, i, 6, result.get(0).getMaHoaDon());
//            }else{
//                if(expResult.equals("null")){
//                    assertEquals(true, true);
//                }
//                Tool.setValue(wb, 0, i, 6,"null");
//            }
//            // TODO review the generated test code and remove the default call to fail.
//            //fail("The test case is a prototype.");
//            Tool.writeEx(wb, filePath);
//        }
//    }

}
