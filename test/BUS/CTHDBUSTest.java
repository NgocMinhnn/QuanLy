/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package BUS;

import DTO.CTHoaDonDTO;
import java.io.IOException;
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
public class CTHDBUSTest {
    String filePath = "CTHDBUS.xlsx";
    public CTHDBUSTest() {
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
     * Test of getData method, of class CTHDBUS.
     */
    @Test
    public void testGetData() {
        System.out.println("getData");
        CTHDBUS instance = new CTHDBUS();
        instance.getData();
        if(instance.list!=null){
            assertEquals(true, true);
        }else
        // TODO review the generated test code and remove the default call to fail.
        fail("Can't getData");
    }

    /**
     * Test of Seacrh method, of class CTHDBUS.
     */
    @Test
    public void testSeacrh() {
        System.out.println("Seacrh");
        CTHDBUS instance = new CTHDBUS();
        Workbook wb = null;
        
        try {
            wb = new XSSFWorkbook("./TestFile/"+filePath);
        } catch (IOException ex) {
            Logger.getLogger(CTHDBUSTest.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        for(int i= 5;i<Tool.getCellNumber(wb, 0, 1, 4)+5;i++){
            String maHD = Tool.getCellStringValue(wb, 0, i, 1);
            ArrayList<CTHoaDonDTO> result = instance.Seacrh(maHD);
            for(int j = 0 ; j<result.size();j++){
                if(!result.get(j).getMaHoaDon().contains(maHD)){
                    fail("Sai MaHD");
                    Tool.setValue(wb, 0, i, 3, "Fail");
                }
            }
            if(result.size()==0){
                Tool.setValue(wb, 0, i, 3, "null");
            }else
                Tool.setValue(wb, 0, i, 3, maHD);
            assertTrue(true);
        }
        Tool.writeEx(wb, filePath);
        // TODO review the generated test code and remove the default call to fail.
    }
    
}
