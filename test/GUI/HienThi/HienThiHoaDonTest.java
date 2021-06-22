/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package GUI.HienThi;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author Vind
 */
public class HienThiHoaDonTest {
    HienThiHoaDon instance;
    public HienThiHoaDonTest() {
        instance = new HienThiHoaDon();
        instance.setVisible(true);
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
     * Test of refresh method, of class HienThiHoaDon.
     */
    @Test
    public void testChiTiet(){
        System.out.println("Test CHi Tiet");
        instance.btnChiTiet.doClick();
        if(instance.k.isShowing()){
            System.out.println("true");
        }
        //phải confirm thì nó mới chạy
        assertEquals(instance.k.isShowing(), false);
    }
}
