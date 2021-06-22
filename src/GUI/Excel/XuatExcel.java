/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package GUI.Excel;

import java.awt.FileDialog;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

import BUS.BaoHanhBUS;
import BUS.HoaDonBUS;
import BUS.KhachHangBUS;
import BUS.LoaiSPBUS;
import BUS.NhaCungCapBUS;
import BUS.NhanVienBUS;
import BUS.PhieuNhapBUS;
import BUS.QuyenBUS;
import BUS.SanPhamBUS;
import BUS.ThuongHieuBUS;
import BUS.Tool;
import DTO.BaoHanhDTO;
import DTO.HoaDonDTO;
import DTO.KhachHangDTO;
import DTO.LoaiSPDTO;
import DTO.NhaCungCapDTO;
import DTO.NhanVienDTO;
import DTO.PhieuNhapDTO;
import DTO.QuyenDTO;
import DTO.SanPhamDTO;
import DTO.ThuongHieuDTO;

/**
 *
 * @author Admin
 */
public class XuatExcel {

    public static JFrame parent;
    FileDialog fd = new FileDialog(new JFrame(), "Xuáº¥t excel", FileDialog.SAVE);

    public static void setParent(JFrame frame) {
        parent = frame;
    }

    private String getFile() {
        fd.setFile("untitla.xls");
        fd.setVisible(true);
        String url = fd.getDirectory() + fd.getFile();
        if (url.equals("nullnull")) {
            return null;
        } else {
            return url;
        }
    }
public void XuatfileExcelSanpham(){
        
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u sáº£n pháº©m ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("TÃ i khoáº£n");

            SanPhamBUS qlSP = new SanPhamBUS();
            ArrayList<SanPhamDTO> list = qlSP.list;

            int rownum = 0;
            Row row = sheet.createRow(rownum);
            //"STT", "MÃ£ thÆ°Æ¡ng hiá»‡u", "TÃªn thÆ°Æ¡ng hiá»‡u", "MÃ´ táº£"
            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ Sáº£n Pháº©m");
            row.createCell(2, CellType.STRING).setCellValue("TÃªn Sáº£n Pháº©m");
            row.createCell(3, CellType.STRING).setCellValue("MÃ£ Loáº¡i Sáº£n Pháº©m");
            row.createCell(4, CellType.STRING).setCellValue("NÄƒm Sáº£n Xuáº¥t");
            row.createCell(5, CellType.STRING).setCellValue("MÃ£ NhÃ  Cung Cáº¥p");
            row.createCell(6, CellType.STRING).setCellValue("MÃ£ ThÆ°Æ¡ng Hiá»‡u");
            row.createCell(7, CellType.NUMERIC).setCellValue("Ä�Æ¡n GiÃ¡");
            row.createCell(8, CellType.NUMERIC).setCellValue("Sá»‘ LÆ°á»£ng");
            row.createCell(9, CellType.STRING).setCellValue("MÃ´ táº£");
            row.createCell(10, CellType.STRING).setCellValue("Image");

            for (SanPhamDTO th : list) {
                rownum++;
                row = sheet.createRow(rownum);

                String masp = th.getMaSanPham();
                String ten = th.getTenSanPham();
                String maloaisp = th.getMaLoaiSP();
                LocalDate namSX = th.getNamSx().toLocalDate();
                String tenNCC = th.getMaNCC();
                String tenthuonghieu = th.getMaThuongHieu();
                int dongia = th.getDongia();
                int soluong = th.getSoLuong();
                String mota = th.getMota();
                String image = th.getImage();

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(masp);
                row.createCell(2, CellType.STRING).setCellValue(ten);
                row.createCell(3, CellType.STRING).setCellValue(maloaisp);
                row.createCell(4, CellType.STRING).setCellValue(namSX.toString());
                row.createCell(5, CellType.STRING).setCellValue(tenNCC);
                row.createCell(6, CellType.STRING).setCellValue(tenthuonghieu);
                row.createCell(7, CellType.NUMERIC).setCellValue(dongia);
                row.createCell(8, CellType.NUMERIC).setCellValue(soluong);
                row.createCell(9, CellType.STRING).setCellValue(mota);
                row.createCell(10, CellType.STRING).setCellValue(image);
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    // Xuáº¥t file Excel NhÃ  cung cáº¥p    
    public void xuatFileExcelNhaCungCap() {
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u nhÃ  cung cáº¥p ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("NhÃ  cung cáº¥p");

            NhaCungCapBUS qlnccBUS = new NhaCungCapBUS();
            ArrayList<NhaCungCapDTO> list = qlnccBUS.list;

            int rownum = 0;
            Row row = sheet.createRow(rownum);

            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ nhÃ  cung cáº¥p");
            row.createCell(2, CellType.STRING).setCellValue("TÃªn nhÃ  cung cáº¥p");
            row.createCell(3, CellType.STRING).setCellValue("Ä�á»‹a chá»‰");
            row.createCell(4, CellType.STRING).setCellValue("Email");
            for (NhaCungCapDTO ncc : list) {
                rownum++;
                row = sheet.createRow(rownum);

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(ncc.getMaNCC());
                row.createCell(2, CellType.STRING).setCellValue(ncc.getTenNCC());
                row.createCell(3, CellType.STRING).setCellValue(ncc.getDiaChi());
                row.createCell(4, CellType.STRING).setCellValue(ncc.getEmail());
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

    // Xuáº¥t file Excel NhÃ¢n viÃªn
    public void xuatFileExcelNhanVien() {
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u nhÃ¢n viÃªn ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("NhÃ¢n viÃªn");

            NhanVienBUS qlnvBUS = new NhanVienBUS();
            ArrayList<NhanVienDTO> list = qlnvBUS.getDsnv();

            int rownum = 0;
            Row row = sheet.createRow(rownum);

            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ nhÃ¢n viÃªn");
            row.createCell(2, CellType.STRING).setCellValue("Há»� nhÃ¢n viÃªn");
            row.createCell(3, CellType.STRING).setCellValue("TÃªn nhÃ¢n viÃªn");
            row.createCell(4, CellType.STRING).setCellValue("NgÃ y sinh");
            row.createCell(5, CellType.STRING).setCellValue("Giá»›i tÃ­nh");
            row.createCell(6, CellType.STRING).setCellValue("Sá»‘ Ä‘iá»‡n thoáº¡i");
            row.createCell(7, CellType.STRING).setCellValue("MÃ£ quyá»�n");
            row.createCell(8, CellType.STRING).setCellValue("LÆ°Æ¡ng");
            row.createCell(9, CellType.STRING).setCellValue("Tráº¡ng thÃ¡i");
//String maNhanVien, String tenNhanVien, String hoNhanVien, java.sql.Date ngaySinh, String gioiTinh, String sdt, String maQuyen, double luong, int trangThai) {
            for (NhanVienDTO nv : list) {
                rownum++;
                row = sheet.createRow(rownum);

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(nv.getMaNhanVien());
                row.createCell(2, CellType.STRING).setCellValue(nv.getHoNhanVien());
                row.createCell(3, CellType.STRING).setCellValue(nv.getTenNhanVien());
                row.createCell(4, CellType.STRING).setCellValue(nv.getNgaySinh().toString());
                row.createCell(5, CellType.STRING).setCellValue(nv.getGioiTinh());
                row.createCell(6, CellType.STRING).setCellValue(nv.getSdt());
                row.createCell(7, CellType.STRING).setCellValue(nv.getMaQuyen());
                row.createCell(8, CellType.STRING).setCellValue(nv.getLuong());
                row.createCell(9, CellType.STRING).setCellValue((nv.getTrangThai() == 0 ? "áº¨n" : "Hiá»‡n"));
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

    // Xuáº¥t file Excel KhÃ¡ch hÃ ng
    public void xuatFileExcelKhachHang() {
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u khÃ¡ch hÃ ng ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("KhÃ¡ch hÃ ng");

            KhachHangBUS qlkhBUS = new KhachHangBUS();
            ArrayList<KhachHangDTO> list = qlkhBUS.list;

            int rownum = 0;
            Row row = sheet.createRow(rownum);
//"MÃ£ khÃ¡ch hÃ ng", "Há»� khÃ¡ch hÃ ng", "TÃªn khÃ¡ch hÃ ng", "NgÃ y sinh", "Sá»‘ Ä‘iá»‡n thoáº¡i", "Loáº¡i KhÃ¡ch HÃ ng", "TÃ­ch LÅ©y"
            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ khÃ¡ch hÃ ng");
            row.createCell(2, CellType.STRING).setCellValue("Há»� khÃ¡ch hÃ ng");
            row.createCell(3, CellType.STRING).setCellValue("TÃªn khÃ¡ch hÃ ng");
            row.createCell(4, CellType.STRING).setCellValue("NgÃ y sinh");
            row.createCell(5, CellType.STRING).setCellValue("Sá»‘ Ä‘iá»‡n thoáº¡i");
            row.createCell(6, CellType.STRING).setCellValue("Loáº¡i khÃ¡ch hÃ ng");
            row.createCell(7, CellType.STRING).setCellValue("TÃ­ch lÅ©y");

            for (KhachHangDTO kh : list) {
                rownum++;
                row = sheet.createRow(rownum);

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(kh.getMaKhachHang());
                row.createCell(2, CellType.STRING).setCellValue(kh.getHoKhachHang());
                row.createCell(3, CellType.STRING).setCellValue(kh.getTenKhachHang());
                row.createCell(4, CellType.STRING).setCellValue(kh.getNgaySinh().toLocalDate().toString());
                row.createCell(5, CellType.STRING).setCellValue(kh.getSdt());
                row.createCell(6, CellType.STRING).setCellValue(kh.getLoaiKhachHang());
                row.createCell(7, CellType.NUMERIC).setCellValue(kh.getTichLuy());
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
//
     public void xuatFileExcelBaoHanh(){
        fd.setTitle("xuáº¥t dá»¯ liá»‡u ra file excel");
        String url=getFile();
        if(url==null){
            return;
            }
           FileOutputStream outFile = null;
           try{
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Báº£o HÃ nh");
        HoaDonBUS qlHD = new HoaDonBUS();
        KhachHangBUS qlKH = new KhachHangBUS();
        SanPhamBUS qlSanPham = new SanPhamBUS();
        LoaiSPBUS qlL = new LoaiSPBUS();
        BaoHanhBUS qlbhBUS =new BaoHanhBUS();
        ArrayList<BaoHanhDTO> list = qlbhBUS.list;
        int rownum = 0;
        Row row = sheet.createRow(rownum);
        //"MÃ£ khÃ¡ch hÃ ng", "H? khÃ¡ch hÃ ng", "TÃªn khÃ¡ch hÃ ng", "NgÃ y sinh", "S? di?n tho?i", "Lo?i KhÃ¡ch HÃ ng", "TÃ­ch Luy"
        row.createCell(0, CellType.NUMERIC).setCellValue("STT");
        row.createCell(1, CellType.STRING).setCellValue("MÃ£ hÃ³a don");
        row.createCell(2, CellType.STRING).setCellValue("TÃªn khÃ¡ch hÃ ng");
        row.createCell(3, CellType.STRING).setCellValue("TÃªn sáº£n pháº©m");
        row.createCell(4, CellType.STRING).setCellValue("Loáº¡i sáº£n pháº©m");
        row.createCell(5, CellType.STRING).setCellValue("NgÃ y láº­p");
        row.createCell(6, CellType.STRING).setCellValue("Thá»�i háº¡n");
        for (BaoHanhDTO bh : list) {
            rownum++;
            row = sheet.createRow(rownum);
            HoaDonDTO hd = qlHD.getHD(bh.getMaHoaDon());
            KhachHangDTO kh = qlKH.getKH(hd.getMaKhachHang());
            SanPhamDTO sp = qlSanPham.getSanPham(bh.getMaSanPham());
            LoaiSPDTO lsp = qlL.getLoaiSPDTO(sp.getMaLoaiSP());
            java.sql.Date a = (java.sql.Date) bh.getNgayLap().clone();
            a.setYear(a.getYear()+2);
            row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
            row.createCell(1, CellType.STRING).setCellValue(bh.getMaHoaDon());
            row.createCell(2, CellType.STRING).setCellValue(kh.getFullName());
            row.createCell(3, CellType.STRING).setCellValue( sp.getTenSanPham());
            row.createCell(4, CellType.STRING).setCellValue( lsp.getTenLoaiSP());
            row.createCell(5, CellType.STRING).setCellValue( bh.getNgayLap().toLocalDate().toString());
            row.createCell(6, CellType.STRING).setCellValue(a.toLocalDate().toString());
            
        }
        for (int i = 0; i < rownum; i++) {
            sheet.autoSizeColumn(i);
        }
        File file = new File(url);
        file.getParentFile().mkdirs();
        outFile = new FileOutputStream(file);
        workbook.write(outFile);
        JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());
        
    }catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
//    // Xuáº¥t file Excel TÃ i khoáº£n
public void xuatFileExcelHoaDon() {
        fd.setTitle("Xuáº¥t hÃ³a Ä‘Æ¡n ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("HÃ³a Ä‘Æ¡n");

            HoaDonBUS qlhdBUS = new HoaDonBUS();
            ArrayList<HoaDonDTO> list = qlhdBUS.list;
// khai bÃ¡o khÃ¡ch hÃ ng, nhÃ¢n viÃªn BUS
            KhachHangBUS qlkh = new KhachHangBUS();
            NhanVienBUS qlnv = new NhanVienBUS();
            int rownum = 0;
            Row row = sheet.createRow(rownum);

            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ hÃ³a Ä‘Æ¡n");
            row.createCell(2, CellType.STRING).setCellValue("TÃªn nhÃ¢n viÃªn");
            row.createCell(3, CellType.STRING).setCellValue("TÃªn khÃ¡ch hÃ ng");
            row.createCell(4, CellType.STRING).setCellValue("Tá»•ng tiá»�n");
            row.createCell(5, CellType.STRING).setCellValue("NgÃ y láº­p");
            row.createCell(6, CellType.STRING).setCellValue("MÃ£ Khuyáº¿n mÃ£i");
            row.createCell(7, CellType.STRING).setCellValue("Giáº£m giÃ¡");
//String maNhanVien, String tenNhanVien, String hoNhanVien, java.sql.Date ngaySinh, String gioiTinh, String sdt, String maQuyen, double luong, int trangThai) {
            for (HoaDonDTO hd : list) {
                rownum++;
                row = sheet.createRow(rownum);

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(hd.getMaHoaDon());
                row.createCell(2, CellType.STRING).setCellValue(qlnv.getNV(hd.getMaNhanVien()).getFullFame());
                row.createCell(3, CellType.STRING).setCellValue(qlkh.getKH(hd.getMaKhachHang()).getFullName());
                row.createCell(4, CellType.NUMERIC).setCellValue(hd.getTongTien());
                row.createCell(5, CellType.STRING).setCellValue(hd.getNgayLap().toString());
                row.createCell(6, CellType.STRING).setCellValue(hd.getMaKM());
                row.createCell(7, CellType.NUMERIC).setCellValue(hd.getGiamGia());
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    public void xuatFileExcelThuongHieu() {
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u thÆ°Æ¡ng hiá»‡u ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("TÃ i khoáº£n");

            ThuongHieuBUS qlTH = new ThuongHieuBUS();
            ArrayList<ThuongHieuDTO> list = qlTH.getTH();

            int rownum = 0;
            Row row = sheet.createRow(rownum);
            //"STT", "MÃ£ thÆ°Æ¡ng hiá»‡u", "TÃªn thÆ°Æ¡ng hiá»‡u", "MÃ´ táº£"
            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ thÆ°Æ¡ng hiá»‡u");
            row.createCell(2, CellType.STRING).setCellValue("TÃªn thÆ°Æ¡ng hiá»‡u");
            row.createCell(3, CellType.STRING).setCellValue("MÃ´ táº£");

            for (ThuongHieuDTO th : list) {
                rownum++;
                row = sheet.createRow(rownum);

                String math = th.getMaThuongHieu();
                String ten = th.getTenThuongHieu();
                String mota = th.getMoTa();

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(math);
                row.createCell(2, CellType.STRING).setCellValue(ten);
                row.createCell(3, CellType.STRING).setCellValue(mota);
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
//
//    // Xuáº¥t file Excel Khuyáº¿n mÃ£i
//    public void xuatFileExcelKhuyenMai() {
//        fd.setTitle("Xuáº¥t dá»¯ liá»‡u khuyáº¿n mÃ£i ra excel");
//        String url = getFile();
//        if (url == null) {
//            return;
//        }
//
//        FileOutputStream outFile = null;
//        try {
//            HSSFWorkbook workbook = new HSSFWorkbook();
//            HSSFSheet sheet = workbook.createSheet("Khuyáº¿n mÃ£i");
//
//            QuanLyKhuyenMaiBUS qlkmBUS = new QuanLyKhuyenMaiBUS();
//            ArrayList<KhuyenMai> list = qlkmBUS.getDskm();
//
//            int rownum = 0;
//            Row row = sheet.createRow(rownum);
//
//            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
//            row.createCell(1, CellType.STRING).setCellValue("MÃ£ khuyáº¿n mÃ£i");
//            row.createCell(2, CellType.STRING).setCellValue("TÃªn khuyáº¿n mÃ£i");
//            row.createCell(3, CellType.NUMERIC).setCellValue("Ä�iá»�u kiá»‡n");
//            row.createCell(4, CellType.NUMERIC).setCellValue("Pháº§n trÄƒm");
//            row.createCell(5, CellType.STRING).setCellValue("NgÃ y báº¯t Ä‘áº§u");
//            row.createCell(6, CellType.STRING).setCellValue("NgÃ y káº¿t thÃºc");
//
//            for (KhuyenMai km : list) {
//                rownum++;
//                row = sheet.createRow(rownum);
//
//                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
//                row.createCell(1, CellType.STRING).setCellValue(km.getMaKM());
//                row.createCell(2, CellType.STRING).setCellValue(km.getTenKM());
//                row.createCell(3, CellType.NUMERIC).setCellValue(km.getDieuKienKM());
//                row.createCell(4, CellType.NUMERIC).setCellValue(km.getPhanTramKM());
//                row.createCell(5, CellType.STRING).setCellValue(String.valueOf(km.getNgayBD()));
//                row.createCell(6, CellType.STRING).setCellValue(String.valueOf(km.getNgayKT()));
//            }
//            for (int i = 0; i < rownum; i++) {
//                sheet.autoSizeColumn(i);
//            }
//
//            File file = new File(url);
//            file.getParentFile().mkdirs();
//            outFile = new FileOutputStream(file);
//            workbook.write(outFile);
//
//            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());
//
//        } catch (FileNotFoundException ex) {
//            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//        } catch (IOException ex) {
//            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//        } finally {
//            try {
//                if (outFile != null) {
//                    outFile.close();
//                }
//            } catch (IOException ex) {
//                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//            }
//        }
//    }
//

    public void xuatFileExcelSanPham() {
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u sáº£n pháº©m ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Sáº£n pháº©m");

            SanPhamBUS qlspBUS = new SanPhamBUS();
            LoaiSPBUS qllsp = new LoaiSPBUS();
            ArrayList<SanPhamDTO> list = qlspBUS.list;

            int rownum = 0;
            Row row = sheet.createRow(rownum);
//            "MÃ£ sáº£n pháº©m", "TÃªn sáº£n pháº©m", "Loáº¡i sáº£n pháº©m", "NÄƒm sx", "NhÃ  cung cáº¥p", "ThÆ°Æ¡ng hiá»‡u", "Ä�Æ¡n giÃ¡", "Sá»‘ lÆ°á»£ng","MÃ´ táº£", "áº¢nh"});
            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ sáº£n pháº©m");
            row.createCell(2, CellType.STRING).setCellValue("TÃªn sáº£n pháº©m");
            row.createCell(3, CellType.STRING).setCellValue("Loáº¡i sáº£n pháº©m");
            row.createCell(4, CellType.STRING).setCellValue("NhÃ  cung cáº¥p");
            row.createCell(5, CellType.STRING).setCellValue("ThÆ°Æ¡ng hiá»‡u");
            row.createCell(6, CellType.NUMERIC).setCellValue("Ä�Æ¡n giÃ¡");
            row.createCell(7, CellType.NUMERIC).setCellValue("Sá»‘ lÆ°á»£ng");
            row.createCell(8, CellType.STRING).setCellValue("MÃ´ táº£");
            row.createCell(9, CellType.STRING).setCellValue("áº¢nh");

            for (SanPhamDTO sp : list) {
                rownum++;
                row = sheet.createRow(rownum);
                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(sp.getMaSanPham());
                row.createCell(2, CellType.STRING).setCellValue(sp.getTenSanPham());
                row.createCell(3, CellType.STRING).setCellValue(sp.getTenLoaiSP() + "-" + sp.getMaLoaiSP());
                row.createCell(4, CellType.STRING).setCellValue(sp.getTenNCC() + "-" + sp.getMaNCC());
                row.createCell(5, CellType.STRING).setCellValue(sp.getTenThuongHieu() + "-" + sp.getMaThuongHieu());
                row.createCell(6, CellType.NUMERIC).setCellValue(sp.getDongia());
                row.createCell(7, CellType.NUMERIC).setCellValue(sp.getSoLuong());
                row.createCell(8, CellType.STRING).setCellValue(sp.getMota());
                row.createCell(9, CellType.STRING).setCellValue(sp.getImage());

            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

//    // Xuáº¥t file Excel Loáº¡i sáº£n pháº©m
    public void xuatFileExcelLoaiSanPham() {
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u loáº¡i sáº£n pháº©m ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Loáº¡i sáº£n pháº©m");

            LoaiSPBUS qllspBUS = new LoaiSPBUS();
            ArrayList<LoaiSPDTO> list = qllspBUS.getds();

            int rownum = 0;
            Row row = sheet.createRow(rownum);

            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ Loáº¡i");
            row.createCell(2, CellType.STRING).setCellValue("TÃªn loáº¡i");

            for (LoaiSPDTO lsp : list) {
                rownum++;
                row = sheet.createRow(rownum);

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(lsp.getMaLoaiSP());
                row.createCell(2, CellType.STRING).setCellValue(lsp.getTenLoaiSP());
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
//xuáº¥t file excel Phiáº¿u nháº­p

    public void xuatFileExcelQuyen() {
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u quyá»�n ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Quyá»�n");

            QuyenBUS qlqBUS = new QuyenBUS();
            ArrayList<QuyenDTO> list = qlqBUS.list;

            int rownum = 0;
            Row row = sheet.createRow(rownum);

            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ quyá»�n");
            row.createCell(2, CellType.STRING).setCellValue("TÃªn quyá»�n");
            row.createCell(3, CellType.STRING).setCellValue("Chi tiáº¿t quyá»�n");

            for (QuyenDTO q : list) {
                rownum++;
                row = sheet.createRow(rownum);

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(q.getMaQuyen());
                row.createCell(2, CellType.STRING).setCellValue(q.getTenQuyen());
                row.createCell(3, CellType.STRING).setCellValue(q.getChitiet());
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

    // Xuáº¥t file Excel Quyá»�n
    public void xuatFileExcelPhieuNhap() {
        fd.setTitle("Xuáº¥t dá»¯ liá»‡u quyá»�n ra excel");
        String url = getFile();
        if (url == null) {
            return;
        }

        FileOutputStream outFile = null;
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Quyá»�n");

            PhieuNhapBUS qlPhieuNhap = new PhieuNhapBUS();
            NhanVienBUS qlNhanVien = new NhanVienBUS();
            //H
            //KhachHangBus qlKH  = new KhachHangBUS();
            //qlKH.getKh()
            NhaCungCapBUS qlNCC = new NhaCungCapBUS();
            ArrayList<PhieuNhapDTO> list = qlPhieuNhap.list;
//mtb.setHeaders(new String[]{"STT", "MÃ£ Phiáº¿u nháº­p", "TÃªn nhÃ¢n viÃªn"," NhÃ  Cung Cáº¥p", "Tá»•ng tiá»�n", "NgÃ y láº­p"});
            int rownum = 0;
            Row row = sheet.createRow(rownum);

            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
            row.createCell(1, CellType.STRING).setCellValue("MÃ£ phiáº¿u nháº­p");
            row.createCell(2, CellType.STRING).setCellValue("tÃªn nhÃ¢n viÃªn");
            row.createCell(3, CellType.STRING).setCellValue("Chi nhÃ  cung cáº¥p");
            row.createCell(4, CellType.STRING).setCellValue("Tá»•ng tiá»�n");
            row.createCell(5, CellType.STRING).setCellValue("NgÃ y láº­p");
//hd.getMaPhieuNhap(),
//                    qlNhanVien.getNV(hd.getMaNhanVien()).getFullFame(),
//                    qlNCC.getNCC(hd.getMaNCC()).getTenNCC(),
//                    Tool.getMonney((int) hd.getTongTien())+",000",
//                    hd.getNgayLap().toString()
            for (PhieuNhapDTO q : list) {
                rownum++;
                row = sheet.createRow(rownum);

                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
                row.createCell(1, CellType.STRING).setCellValue(q.getMaPhieuNhap());
                row.createCell(2, CellType.STRING).setCellValue(qlNhanVien.getNV(q.getMaNhanVien()).getFullFame());
                row.createCell(3, CellType.STRING).setCellValue(qlNCC.getNCC(q.getMaNCC()).getTenNCC());
                row.createCell(4, CellType.NUMERIC).setCellValue(Tool.getMonney((int) q.getTongTien()) + ",000");
                row.createCell(5, CellType.STRING).setCellValue(q.getNgayLap().toString());
            }
            for (int i = 0; i < rownum; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(url);
            file.getParentFile().mkdirs();
            outFile = new FileOutputStream(file);
            workbook.write(outFile);

            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());

        } catch (FileNotFoundException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                if (outFile != null) {
                    outFile.close();
                }
            } catch (IOException ex) {
                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

//    // Xuáº¥t file Excel HÃ³a Ä‘Æ¡n
//    public void xuatFileExcelHoaDon() {
//        fd.setTitle("Xuáº¥t dá»¯ liá»‡u hÃ³a Ä‘Æ¡n ra excel");
//        String url = getFile();
//        if (url == null) {
//            return;
//        }
//
//        FileOutputStream outFile = null;
//        try {
//            HSSFWorkbook workbook = new HSSFWorkbook();
//            HSSFSheet sheet = workbook.createSheet("HÃ³a Ä‘Æ¡n");
//
//            QuanLyHoaDonBUS qlhdBUS = new QuanLyHoaDonBUS();
//            QuanLyChiTietHoaDonBUS qlcthdBUS = new QuanLyChiTietHoaDonBUS();
//            QuanLyNhanVienBUS qlnvBUS = new QuanLyNhanVienBUS();
//            QuanLyKhachHangBUS qlkhBUS = new QuanLyKhachHangBUS();
//            QuanLyKhuyenMaiBUS qlkmBUS = new QuanLyKhuyenMaiBUS();
//            QuanLySanPhamBUS qlspBUS = new QuanLySanPhamBUS();
//            ArrayList<HoaDon> list = qlhdBUS.getDshd();
//
//            int rownum = 0;
//            int sttHoaDon = 0;
//            Row row = sheet.createRow(rownum);
//
//            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
//            row.createCell(1, CellType.STRING).setCellValue("MÃ£ hÃ³a Ä‘Æ¡n");
//            row.createCell(2, CellType.STRING).setCellValue("MÃ£ nhÃ¢n viÃªn");
//            row.createCell(3, CellType.STRING).setCellValue("MÃ£ khÃ¡ch hÃ ng");
//            row.createCell(4, CellType.STRING).setCellValue("MÃ£ khuyáº¿n mÃ£i");
//            row.createCell(5, CellType.STRING).setCellValue("NgÃ y láº­p");
//            row.createCell(6, CellType.STRING).setCellValue("Giá»� láº­p");
//            row.createCell(7, CellType.STRING).setCellValue("Tá»•ng tiá»�n");
//
//            row.createCell(8, CellType.STRING).setCellValue("Sáº£n pháº©m");
//            row.createCell(9, CellType.STRING).setCellValue("Sá»‘ lÆ°á»£ng");
//            row.createCell(10, CellType.STRING).setCellValue("Ä�Æ¡n giÃ¡");
//            row.createCell(11, CellType.STRING).setCellValue("ThÃ nh tiá»�n");
//
//            for (HoaDon hd : list) {
//                rownum++;
//                sttHoaDon++;
//                row = sheet.createRow(rownum);
//
//                row.createCell(0, CellType.NUMERIC).setCellValue(sttHoaDon);
//                row.createCell(1, CellType.STRING).setCellValue(hd.getMaHoaDon());
//                row.createCell(2, CellType.STRING).setCellValue(hd.getMaNhanVien() + " - " + qlnvBUS.getNhanVien(hd.getMaNhanVien()).getTenNV());
//                row.createCell(3, CellType.STRING).setCellValue(hd.getMaKhachHang() + " - " + qlkhBUS.getKhachHang(hd.getMaKhachHang()).getTenKH());
//                row.createCell(4, CellType.STRING).setCellValue(hd.getMaKhuyenMai() + " - " + qlkmBUS.getKhuyenMai(hd.getMaKhuyenMai()).getTenKM());
//                row.createCell(5, CellType.STRING).setCellValue(String.valueOf(hd.getNgayLap()));
//                row.createCell(6, CellType.STRING).setCellValue(String.valueOf(hd.getGioLap()));
//                row.createCell(7, CellType.NUMERIC).setCellValue(hd.getTongTien());
//
//                for (ChiTietHoaDon cthd : qlcthdBUS.getAllChiTiet(hd.getMaHoaDon())) {
//                    rownum++;
//                    row = sheet.createRow(rownum);
//
//                    String masp = cthd.getMaSanPham();
//
//                    row.createCell(8, CellType.STRING).setCellValue(masp + " - " + qlspBUS.getSanPham(masp).getTenSP());
//                    row.createCell(9, CellType.NUMERIC).setCellValue(cthd.getSoLuong());
//                    row.createCell(10, CellType.NUMERIC).setCellValue(cthd.getDonGia());
//                    row.createCell(11, CellType.NUMERIC).setCellValue(cthd.getDonGia() * cthd.getSoLuong());
//                }
//            }
//            for (int i = 0; i < rownum; i++) {
//                sheet.autoSizeColumn(i);
//            }
//
//            File file = new File(url);
//            file.getParentFile().mkdirs();
//            outFile = new FileOutputStream(file);
//            workbook.write(outFile);
//
//            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());
//
//        } catch (FileNotFoundException ex) {
//            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//        } catch (IOException ex) {
//            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//        } finally {
//            try {
//                if (outFile != null) {
//                    outFile.close();
//                }
//            } catch (IOException ex) {
//                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//            }
//        }
//    }
//
//    // Xuáº¥t file Excel Phiáº¿u nháº­p
//    public void xuatFileExcelPhieuNhap() {
//        fd.setTitle("Xuáº¥t dá»¯ liá»‡u phiáº¿u nháº­p ra excel");
//        String url = getFile();
//        if (url == null) {
//            return;
//        }
//
//        FileOutputStream outFile = null;
//        try {
//            HSSFWorkbook workbook = new HSSFWorkbook();
//            HSSFSheet sheet = workbook.createSheet("HÃ³a Ä‘Æ¡n");
//
//            QuanLyPhieuNhapBUS qlpnBUS = new QuanLyPhieuNhapBUS();
//            QuanLyChiTietPhieuNhapBUS qlctpnBUS = new QuanLyChiTietPhieuNhapBUS();
//            QuanLySanPhamBUS qlspBUS = new QuanLySanPhamBUS();
//            QuanLyNhanVienBUS qlnvBUS = new QuanLyNhanVienBUS();
//            QuanLyNhaCungCapBUS qlnccBUS = new QuanLyNhaCungCapBUS();
//            ArrayList<PhieuNhap> list = qlpnBUS.getDspn();
//
//            int rownum = 0;
//            Row row = sheet.createRow(rownum);
//
//            row.createCell(0, CellType.NUMERIC).setCellValue("STT");
//            row.createCell(1, CellType.STRING).setCellValue("MÃ£ hÃ³a Ä‘Æ¡n");
//            row.createCell(2, CellType.STRING).setCellValue("MÃ£ nhÃ  cung cáº¥p");
//            row.createCell(3, CellType.STRING).setCellValue("MÃ£ nhÃ¢n viÃªn");
//            row.createCell(4, CellType.STRING).setCellValue("NgÃ y láº­p");
//            row.createCell(5, CellType.STRING).setCellValue("Giá»� láº­p");
//            row.createCell(6, CellType.STRING).setCellValue("Tá»•ng tiá»�n");
//            row.createCell(7, CellType.STRING).setCellValue("Sáº£n pháº©m");
//            row.createCell(8, CellType.STRING).setCellValue("Sá»‘ lÆ°á»£ng");
//            row.createCell(9, CellType.STRING).setCellValue("Ä�Æ¡n giÃ¡");
//            row.createCell(10, CellType.STRING).setCellValue("ThÃ nh tiá»�n");
//
//            for (PhieuNhap pn : list) {
//                rownum++;
//                row = sheet.createRow(rownum);
//
//                row.createCell(0, CellType.NUMERIC).setCellValue(rownum);
//                row.createCell(1, CellType.STRING).setCellValue(pn.getMaPN());
//                row.createCell(2, CellType.STRING).setCellValue(pn.getMaNCC() + " - " + qlnccBUS.getNhaCungCap(pn.getMaNCC()).getTenNCC());
//                row.createCell(3, CellType.STRING).setCellValue(pn.getMaNV() + " - " + qlnvBUS.getNhanVien(pn.getMaNV()).getTenNV());
//                row.createCell(4, CellType.STRING).setCellValue(String.valueOf(pn.getNgayNhap()));
//                row.createCell(5, CellType.STRING).setCellValue(String.valueOf(pn.getGioNhap()));
//                row.createCell(6, CellType.NUMERIC).setCellValue(pn.getTongTien());
//
//                for (ChiTietPhieuNhap ctpn : qlctpnBUS.getAllChiTiet(pn.getMaPN())) {
//                    rownum++;
//                    row = sheet.createRow(rownum);
//
//                    String masp = ctpn.getMaSP();
//
//                    row.createCell(7, CellType.STRING).setCellValue(masp + " - " + qlspBUS.getSanPham(masp).getTenSP());
//                    row.createCell(8, CellType.NUMERIC).setCellValue(ctpn.getSoLuong());
//                    row.createCell(9, CellType.NUMERIC).setCellValue(ctpn.getDonGia());
//                    row.createCell(10, CellType.NUMERIC).setCellValue(ctpn.getDonGia() * ctpn.getSoLuong());
//                }
//            }
//            for (int i = 0; i < rownum; i++) {
//                sheet.autoSizeColumn(i);
//            }
//
//            File file = new File(url);
//            file.getParentFile().mkdirs();
//            outFile = new FileOutputStream(file);
//            workbook.write(outFile);
//
//            JOptionPane.showMessageDialog(null, "Ghi file thÃ nh cÃ´ng: " + file.getAbsolutePath());
//
//        } catch (FileNotFoundException ex) {
//            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//        } catch (IOException ex) {
//            Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//        } finally {
//            try {
//                if (outFile != null) {
//                    outFile.close();
//                }
//            } catch (IOException ex) {
//                Logger.getLogger(XuatExcel.class.getName()).log(Level.SEVERE, null, ex);
//            }
//        }
//    }
    private String getTime() {
        return new SimpleDateFormat("yyyyMMddHHmm").format(new Date());
    }
}
