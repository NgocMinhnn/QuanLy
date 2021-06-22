/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package BUS;

import DTO.HoaDonDTO;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.annotation.Repeatable;
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
public class HoaDonBUSTest {
    static String filePath = "HoaDonBUS.xlsx";
    Workbook wb = null;
    public HoaDonBUSTest() {
        
    }
    
    @BeforeClass
    public static void setUpClass() {
    }
    
    @AfterClass
    public static void tearDownClass() {
    }
    
    @Before
    public void setUp() {
        try{
            FileInputStream file = new FileInputStream("./TestFile/"+filePath);
            wb = new XSSFWorkbook(file);
        }catch(Exception e){
            e.printStackTrace();
        }
    }
    
    @After
    public void tearDown() {
        Tool.writeEx(wb, filePath);
    }

    /**
     * Test of bigNum method, of class HoaDonBUS.
     */
    @Test
    public void testBigNum() {
        System.out.println("bigNum");
        HoaDonBUS instance = new HoaDonBUS();  
        String expResult = "13";
        String result = instance.bigNum();
        assertEquals(expResult, result);
        Tool.setValue(wb, 1, 5, 3, result);
        //Tool.writeEx(wb, filePath);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }
//
//    /**
//     * Test of getNextMaHD method, of class HoaDonBUS.
//     */
    @Test
    public void testGetNextMaHD() {
        System.out.println("getNextMaHD");
        HoaDonBUS instance = new HoaDonBUS();
        String expResult = Tool.getCellStringValue(wb, 0, 5, 2);
        String result = instance.getNextMaHD();
        assertEquals(expResult, result);
        Tool.setValue(wb, 0, 5, 3, result);
        //Tool.writeEx(wb, filePath);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }
//
//    /**
//     * Test of getHD method, of class HoaDonBUS.
//     */
    @Test
    public void testGetHD() {
        System.out.println("getHD");
        HoaDonBUS instance = new HoaDonBUS();
        for(int i=5;i<Tool.getCellNumber(wb, 2, 1, 5)+5;i++){
            String mahd = Tool.getCellStringValue(wb, 2, i, 1);
            String expResult = Tool.getCellStringValue(wb, 2, i, 2);
            HoaDonDTO result = instance.getHD(mahd);
            if(result==null){
                assertEquals(null, result);
                Tool.setValue(wb, 2, i, 3, new String("null"));
            }else{
                assertEquals(expResult, result.getMaHoaDon());
                Tool.setValue(wb, 2, i, 3, result.getMaHoaDon());
            }
        }
        //Tool.writeEx(wb, filePath);
    }
//
//    /**
//     * Test of getData method, of class HoaDonBUS.
//     */
    @Test
    public void testGetData() {
        System.out.println("getData");
        HoaDonBUS instance = new HoaDonBUS();
        instance.getData();
        if(instance.list==null){
            Tool.setValue(wb, 3, 5, 3, "null");
        }else{
            Tool.setValue(wb, 3, 5, 3, "list!=null");
        }
        //Tool.writeEx(wb, filePath);
        // 0TODO review the generated test code and remove the default call to fail.
        assertFalse(instance.list.equals(null));
    }
//
//    /**
//     * Test of Search method, of class HoaDonBUS.
//     */
    @Test
    public void testSearch() {
        //Tool.writeEx(wb, filePath);
        System.out.println("Search");
        String value = "";
        String type = "";
        LocalDate date1 = null;
        LocalDate date2 = null;
        int gia1 = 0;
        int gia2 = 0;
        HoaDonBUS instance = new HoaDonBUS();
        ArrayList<HoaDonDTO> result;
        for(int i=5;i<Tool.getCellNumber(wb, 4, 1, 10)+5;i++){
            value = Tool.getCellStringValue(wb, 4, i, 1);
            type = Tool.getCellStringValue(wb, 4, i, 2);
            date1= Tool.getCellDate(wb, 4, i, 3);
            date2= Tool.getCellDate(wb, 4, i, 4);
            gia1 = Tool.getInt(Tool.getCellStringValue(wb, 4, i, 5));
            gia2 = Tool.getInt(Tool.getCellStringValue(wb, 4, i, 6));
            result = instance.Search(value, type, date1, date2, gia1, gia2);
            String EO = Tool.getCellStringValue(wb, 4, i, 7);
            //assertEquals(EO, result.size()+"");
            Tool.setValue(wb, 4, i, 8, result.size()+"");
        }
        
        //assertEquals(expResult, result);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }   
}
