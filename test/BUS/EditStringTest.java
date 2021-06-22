/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package BUS;

import java.io.IOException;
import java.sql.Date;
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
public class EditStringTest {
    String filePath="EditString.xlsx";
    public EditStringTest() {
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
     * Test of DateToString method, of class EditString.
     */
    @Test
    public void testDateToString() {
        System.out.println("DateToString");
        Date date = null;
        Workbook wb = null;
        
        try {
            wb = new XSSFWorkbook("./TestFile/"+filePath);
        } catch (IOException ex) {
            Logger.getLogger(CTHDBUSTest.class.getName()).log(Level.SEVERE, null, ex);
        }
         for(int i=5;i<Tool.getCellNumber(wb, 0, 1, 4)+4;i++){
            String expResult =Tool.getCellStringValue(wb, 0, i, 2);
            date = Date.valueOf(Tool.getCellDate(wb, 0, i, 1));
            String result = EditString.DateToString(date);
            System.out.println(result);
            Tool.setValue(wb, 0, i, 3, result);
            //assertEquals(expResult, result);     
         }
        Tool.writeEx(wb, filePath);
        // TODO review the generated test code and remove the default call to fail.
        //fail("The test case is a prototype.");
    }
    
}
