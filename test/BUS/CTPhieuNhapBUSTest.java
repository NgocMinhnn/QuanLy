/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package BUS;

import DTO.CTPhieuNhapDTO;
import java.util.ArrayList;
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
public class CTPhieuNhapBUSTest {
    String filePath = "CTPhieuNhapBUS.xlsx";
    public CTPhieuNhapBUSTest() {
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
     * Test of getData method, of class CTPhieuNhapBUS.
     */
    @Test
    public void testGetData() {
        System.out.println("getData");
        CTPhieuNhapBUS instance = new CTPhieuNhapBUS();
        Workbook wb = null;
        try{
            wb = new XSSFWorkbook("./TestFile/"+filePath);
        }catch(Exception e){
            e.printStackTrace();
        }
        for(int i= 5;i<Tool.getCellNumber(wb, 0, 1, 4)+5;i++){
            instance.getData();
            if(instance.list!=null){
                assertTrue(true);
                Tool.setValue(wb, 0, i, 3, "!null");
            }else{
                assertTrue(false);
                Tool.setValue(wb, 0, i, 3, "null");
            }
        }
        Tool.writeEx(wb, filePath);
        // TODO review the generated test code and remove the default call to fail.
    }

    /**
     * Test of Seacrh method, of class CTPhieuNhapBUS.
     */
    @Test
    public void testSeacrh() {
        System.out.println("Seacrh");        
        String maPN = "";
        CTPhieuNhapBUS instance = new CTPhieuNhapBUS();
        Workbook wb = null;
        try{
            wb = new XSSFWorkbook("./TestFile/"+filePath);
        }catch(Exception e){
            e.printStackTrace();
        }
        ArrayList<CTPhieuNhapDTO> result;
        for(int i=5;i<Tool.getCellNumber(wb, 1, 1, 4)+5;i++){
            maPN = Tool.getCellStringValue(wb, 1, i, 1);
            result = instance.Seacrh(maPN);
            if(result.size()>0){
                int k=0;
                for(int j=0;j<result.size();j++){
                    if(!result.get(j).getMaPhieuNhap().equals(maPN)){
                        Tool.setValue(wb, 1, i, 3, "Fail");
                        fail("Erro");
                        k++;
                        break;
                    }
                }
                if(k==0){
                    Tool.setValue(wb, 1, i, 3, maPN);
                }
            }else{
                Tool.setValue(wb, 1, i, 3, "null");
            }
        }
        Tool.writeEx(wb, filePath);
    }
    
}
