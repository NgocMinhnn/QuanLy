/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package GUI.Quanli;

import BUS.LoaiSPBUS;
import BUS.ThuongHieuBUS;
import GUI.HienThi.*;
import DAO.LoaiSPDAO;
import DAO.NhanVienDAO;
import DAO.ThuongHieuDAO;
import DTO.LoaiSPDTO;
import DTO.ThuongHieuDTO;
import GUI.Button.*;
import GUI.*;
import GUI.Excel.DocExcel;
import GUI.Excel.XuatExcel;
import GUI.Quanli.ThemSuaThuongHieu;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import javax.swing.*;
import javax.swing.event.CellEditorListener;
import javax.swing.event.ChangeEvent;
import javax.swing.event.TableModelListener;
public class QuanLiLoaiSP extends JPanel {
    HienThiLoaiSP formHienThi = new HienThiLoaiSP();
    ThemButton btnThem = new ThemButton();
    SuaButton btnSua = new SuaButton();
    XoaButton btnXoa = new XoaButton();
    //NhanVienDTO nvSua = new NhanVienDTO();
    ExportExcelButton btnXuatExcel = new ExportExcelButton();
    ImportExcelButton btnNhapExcel = new ImportExcelButton();

    public QuanLiLoaiSP()  {
        setLayout(new BorderLayout());

        // buttons
        if (!DangNhap.quyenLogin.getChitiet().contains("qlLoaiSanPham")) {
            btnThem.setEnabled(false);
            btnXoa.setEnabled(false);
            btnSua.setEnabled(false);
            btnNhapExcel.setEnabled(false);
            btnXuatExcel.setEnabled(false);
            formHienThi.getTable().getTable().setEnabled(false);
        }

        JPanel plBtn = new JPanel();
        plBtn.add(btnThem);
        plBtn.add(btnXoa);
        plBtn.add(btnSua);
        plBtn.add(btnXuatExcel);
        plBtn.add(btnNhapExcel);

        this.add(plBtn, BorderLayout.NORTH);
        this.add(formHienThi, BorderLayout.CENTER);
        //action Listener
        btnThem.addActionListener((ActionEvent ae) -> {
            btnThemMouseClicked();
        });
        btnXoa.addActionListener((ActionEvent ae) -> {
            btnXoaMouseClicked();
        });
        btnSua.addActionListener((ActionEvent ae) -> {
            btnSuaMouseClicked();
        });
        btnXuatExcel.addActionListener((ActionEvent ae) -> {
            new XuatExcel().xuatFileExcelLoaiSanPham();
            
        });
        btnNhapExcel.addActionListener((ActionEvent ae) -> {
            new DocExcel().docFileExcelLoaiSanPham();
        });
    }

    private void btnThemMouseClicked() {
        ThemSuaLoaiSP themLoaiSP = new ThemSuaLoaiSP("Th??m", "");
        themLoaiSP.setVisible(true);
        themLoaiSP.addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosed(java.awt.event.WindowEvent windowEvent) {
                formHienThi.refresh();
            }
        });
    }

    private void btnXoaMouseClicked() {
        if (formHienThi.getTable().getTable().getSelectedRow()>=0) {
            LoaiSPDTO lsp = new LoaiSPBUS().getLoaiSPDTO(formHienThi.getSelectedRow(1));
           int reply = JOptionPane.showConfirmDialog(formHienThi, "B???n c?? mu???n ???n lo???i s???n ph???m ?");
           if ( reply == JOptionPane.YES_OPTION) {
               if(LoaiSPDAO.DeleteMaSP(lsp))
               JOptionPane.showMessageDialog(this, "lo???i s???n ph???m ???? ???????c x??a");
               else{
                   JOptionPane.showMessageDialog(this,"X??a lo???i s???n ph???m kh??ng th??nh kh??ng");
               }
               formHienThi.refresh();
           }else{
               JOptionPane.showMessageDialog(this,"X??a lo???i s???n ph???m kh??ng th??nh kh??ng");
           }
            
        }else{
            JOptionPane.showMessageDialog(formHienThi,"B???n ph???i ch???n 1 lo???i s???n ph???m ????? c?? th??? x??a");
        }
    }

    private void btnSuaMouseClicked() {
        ThemSuaLoaiSP themLsp = new ThemSuaLoaiSP("S???a", formHienThi.getSelectedRow(1));
        themLsp.setVisible(true);
        themLsp.addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosed(java.awt.event.WindowEvent windowEvent) {
                formHienThi.refresh();
            }
        });
    }
}


