package GUI.Quanli;

import BUS.KhachHangBUS;
import DAO.KhachHangDAO;
import DTO.KhachHangDTO;
import GUI.Button.ExportExcelButton;
import GUI.Button.ImportExcelButton;
import GUI.Button.SuaButton;
import GUI.Button.ThemButton;
import GUI.Button.XoaButton;
import GUI.DangNhap;
import GUI.Excel.DocExcel;
import GUI.Excel.XuatExcel;
import GUI.HienThi.HienThiKhachHang;
import GUI.HienThi.HienThiKhachHang;
import GUI.HienThi.ThemSuaKhachHang;
import java.awt.BorderLayout;
import java.awt.event.ActionEvent;
import javax.swing.JOptionPane;
import javax.swing.JPanel;

public class QuanLiKhachHangForm extends JPanel {

    HienThiKhachHang formhienthi = new HienThiKhachHang();
    ThemButton btnthem = new ThemButton();
    SuaButton btnsua = new SuaButton();
    XoaButton btnxoa = new XoaButton();

    ExportExcelButton btnxuatexcel = new ExportExcelButton();
    ImportExcelButton btnnhapexcel = new ImportExcelButton();

    public QuanLiKhachHangForm() {
        setLayout(new BorderLayout());

        //button
        if (!DangNhap.quyenLogin.getChitiet().contains("qlKhachHang")) {
            btnthem.setEnabled(false);
            btnxoa.setEnabled(false);
            btnsua.setEnabled(false);
            btnnhapexcel.setEnabled(false);
            btnxuatexcel.setEnabled(false);
            formhienthi.getTable().getTable().setEnabled(false);
        }

        JPanel plBtn = new JPanel();
        plBtn.add(btnthem);
        plBtn.add(btnxoa);
        plBtn.add(btnsua);
        plBtn.add(btnnhapexcel);
        plBtn.add(btnxuatexcel);

        this.add(plBtn, BorderLayout.NORTH);
        this.add(formhienthi, BorderLayout.CENTER);
        //action Listener
        btnthem.addActionListener((ActionEvent ae) -> {
            btnThemMouseClicked();
        });
        btnxoa.addActionListener((ActionEvent ae) -> {
            btnXoaMouseClicked();
        });
        btnsua.addActionListener((ActionEvent ae) -> {
            btnSuaMouseClicked();
        });
        btnnhapexcel.addActionListener((ActionEvent ae) -> {
            new DocExcel().docFileExcelKhachhang();
        });
        btnxuatexcel.addActionListener((ActionEvent ae) -> {
            new XuatExcel().xuatFileExcelKhachHang();
        });

    }

    private void btnThemMouseClicked() {
        ThemSuaKhachHang themKh = new ThemSuaKhachHang("Th??m", "");
        themKh.setVisible(true);
        themKh.addWindowListener(new java.awt.event.WindowAdapter() {

            @Override
            public void windowClosed(java.awt.event.WindowEvent windowevent) {
                formhienthi.refresh();
            }
        });
    }

    private void btnXoaMouseClicked() {
        KhachHangDTO kh = new KhachHangBUS().getKH(formhienthi.getSelectedRow(1));

        int reply = JOptionPane.showConfirmDialog(formhienthi, "B???n c?? mu???n x??a kh??ch h??ng?????????");
        if (kh != null) {
            if (reply == JOptionPane.YES_OPTION) {
                if(KhachHangDAO.DeleteKhachHang(kh)){
                    JOptionPane.showMessageDialog(formhienthi, "x??a kh??ch h??ng th??nh c??ng");
                }else{
                    JOptionPane.showMessageDialog(formhienthi, "x??a kh??ch h??ng kh??ng th??nh c??ng");
                }
                formhienthi.refresh();
            } else {
                JOptionPane.showMessageDialog(formhienthi, "x??a kh??ch h??ng kh??ng th??nh c??ng");
            }
        } else {
            JOptionPane.showMessageDialog(this, "Ch???n kh??ch h??ng ????? x??a");
        }

    }

    private void btnSuaMouseClicked() {
        ThemSuaKhachHang themnv = new ThemSuaKhachHang("S???a", formhienthi.getSelectedRow(1));
        themnv.setVisible(true);
        themnv.addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosed(java.awt.event.WindowEvent windowEvent) {
                formhienthi.refresh();
            }
        });
    }

}
