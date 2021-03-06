/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package GUI;

import GUI.Excel.XuatExcel;
import GUI.HienThi.ChonNhaCungCap;
import GUI.Quanli.QuanLiKhachHangForm;
import GUI.NavBar.*;
import GUI.Quanli.*;
import com.itextpdf.text.DocumentException;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import static java.awt.Frame.ICONIFIED;
import java.awt.Rectangle;
import javax.swing.BorderFactory;
import javax.swing.JFrame;
import javax.swing.JScrollPane;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.awt.event.MouseMotionAdapter;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.border.Border;

/**
 *
 * @author phuon
 */
public class Main extends JFrame implements MouseListener {

    final int WIDTH = 1920, HEIGHT = 1000;
    int px, py;
    NavBarContainer menu, header;
    NavBarButton currentTab;
    JPanel plContent = new JPanel();
    NavBarTitle headerTitle;
    EmptyPage emptypage = new EmptyPage();
    QuanliNhanVienForm nhanvien;
    BanHangForm banhang;
    NhapHangForm nhaphang;
    QuanLiKhachHangForm khachhang;
    QuanLiSanPhamForm sanpham;
    ThongKe thongke;
    QuanLiNhaCungCap nhacungcap;
    QuanLiKhuyenMaiForm khuyenmai;
    QuanLiThuongHieu thuonghieu;
    QuanLiHoaDon hoadon;
    QuanLiBaoHanh baohanh;
    QuanLiQuyen quyen;
    QuanLiLoaiSP lsanpham;
    QuanLiPhieuNhap phieunhap;
    public static JFrame All = new JFrame();
    public int status;

    //test
    //String quyen = "qlBanHangqlNhapHangqlNCCqlQuyenqlKhachHangqlSanPhamqlLoaiSanPhamqlHoaDonqlPhieuNhapqlKhuyenMai";
    public Main() {
        XuatExcel.setParent(this);
        status = 1;
        setLayout(new BorderLayout());
        setSize(WIDTH, HEIGHT);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setUndecorated(true);
        //setResizable(false);
        //setResizable(true);
        setLocationRelativeTo(null);
        String[] navItemInfo = {
            "seperate", "2", "", "",
            "B??n h??ng", "banhang.png", "qlBanHang", "qlBanHang",
            "Nh???p h??ng", "nhaphang.png", "qlNhapHang", "qlNhapHang",
            "seperate", "2", "", "",
            "S???n ph???m", "sanpham.png", "xemSanPham", "qlSanPham",
            "Lo???i s???n ph???m", "loaisp.png", "xemLoaiSanPham", "qlLoaiSanPham",
            "seperate", "2", "", "",
            "H??a ????n", "hoadon.png", "xemHoaDon", "qlHoaDon",
            "Phi???u nh???p", "truck_50px.png", "xemPheuNhap", "qlPhieuNhap",
            "Khuy???n m??i", "gift_50px.png", "xemKhuyenMai", "qlKhuyenMai",
            "B???o h??nh", "contract_100px.png", "xemBaoHanh", "qlBaoHanh",
            "seperate", "1", "", "",
            "Nh??n vi??n", "support_50px.png", "xemNhanVien", "qlNhanVien",
            "Kh??ch h??ng", "user_male_50px.png", "xemKhachHang", "qlKhachHang",
            "Nh?? cung c???p", "company_50px.png", "xemNCC", "qlNCC",
            "Th????ng hi???u", "company_50px.png", "xemThuongHieu", "qlThuongHieu",
            "seperate", "1", "", "",
            "T??i kho???n", "key_50px.png", "xemTaiKhoan", "qlTaiKhoan",
            "Quy???n", "lock_50px.png", "xemQuyen", "qlQuyen",
            "seperate", "1", "", "",
            "Th???ng k??", "bar_chart_50px.png", "", "",
        };
        int menuW = 250;
        int menuH = 0;
        menu = new NavBarContainer(new Rectangle(0, 0, menuW, HEIGHT));
        for (int i = 0; i < navItemInfo.length; i += 4) {
            if (navItemInfo[i].equals("seperate")) {
                NavBarSeperator s = new NavBarSeperator(new Rectangle(0, 0, 0, Integer.parseInt(navItemInfo[i + 1])));
                menu.addItem(s);
            } else {

                String chitietquyen = DangNhap.quyenLogin.getChitiet();
                if (chitietquyen.contains(navItemInfo[i + 2]) || chitietquyen.contains(navItemInfo[i + 3])) {
                    NavBarButton nb = new NavBarButton(new Rectangle(0, 0, 0, 60), navItemInfo[i], navItemInfo[i + 1]);
                    nb.addMouseListener(this);
                    menu.addItem(nb);
                    menuH += 60;
                }
            }
        }

        //https://stackoverflow.com/questions/1385737/scrollable-jpanel
        //https://stackoverflow.com/questions/5590242/how-to-hide-the-jscrollbars-in-a-jscrollpane
        //https://stackoverflow.com/questions/5583495/how-do-i-speed-up-the-scroll-speed-in-a-jscrollpane-when-using-the-mouse-wheel
        JScrollPane scrollMenu = new JScrollPane(menu, JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED, JScrollPane.HORIZONTAL_SCROLLBAR_NEVER);
        menu.setAutoscrolls(true);
        menu.setPreferredSize(new Dimension(menuW, menuH + 100));
        scrollMenu.setPreferredSize(new Dimension(menuW, HEIGHT));
        scrollMenu.setBorder(BorderFactory.createEmptyBorder());
        scrollMenu.getVerticalScrollBar().setUnitIncrement(5);
// ================ Header ===================
        int headerBg = 30;
        int headerH = 55;
        header = new NavBarContainer(new Rectangle(0, 0, WIDTH, headerH));
        header.setBackground(new Color(headerBg, headerBg, headerBg));

        headerTitle = new NavBarTitle(new Rectangle((WIDTH - 400) / 2, 0, 400, headerH), "QU???N L??");
        headerTitle.setColorDefault(new Color(200, 200, 200));
        headerTitle.setBgDefault(new Color(headerBg, headerBg, headerBg));
        headerTitle.setFontSize(23);
        header.addItem(headerTitle, false);

        // Close Button
        int btnWidth = 50;
        int iconSize = 30;
        NavBarButton btnClose = new NavBarButton(new Rectangle(WIDTH - btnWidth, 0, btnWidth, headerH), "", "cancel_100px.png");
        btnClose.setIconLocation(new Rectangle((btnWidth - iconSize) / 2, (headerH - iconSize) / 2, iconSize, iconSize));
        btnClose.setBgDefault(new Color(headerBg, headerBg, headerBg));
        btnClose.setBgHover(new Color(190, 45, 45));
        btnClose.setToolTipText("Tho??t");
        btnClose.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent me) {
                int reply = JOptionPane.showConfirmDialog(getRootPane(),
                        "B???n c?? ch???c mu???n tho??t ch????ng tr??nh Qu???n L???", "Ch?? ??",
                        JOptionPane.YES_NO_OPTION);
                if (reply == JOptionPane.YES_OPTION) {
                    System.exit(0);
                }
            }
        });
        header.addItem(btnClose, false);

        // Minimize Button
        NavBarButton btnMinimize = new NavBarButton(new Rectangle(WIDTH - btnWidth * 2, 0, btnWidth, headerH), "", "below_50px.png");
        btnMinimize.setIconLocation(new Rectangle((btnWidth - iconSize) / 2, (headerH - iconSize) / 2, iconSize, iconSize));
        btnMinimize.setBgDefault(new Color(headerBg, headerBg, headerBg));
        btnMinimize.setBgHover(new Color(49, 49, 49));
        btnMinimize.setToolTipText("Thu nh???");
        btnMinimize.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent me) {
                setState(ICONIFIED);
            }
        });
        header.addItem(btnMinimize, false);

        // logout button
        if (DangNhap.taiKhoanLogin != null) {

            String tenNhanVien = DangNhap.getTenNV();

            NavBarButton btnLogout = new NavBarButton(new Rectangle(0, 0, menuW - btnWidth, headerH), tenNhanVien, "sign_out_50px.png");
            btnLogout.setBgDefault(new Color(headerBg, headerBg, headerBg));
            btnLogout.setBgHover(new Color(49, 49, 49));
            btnLogout.relocate2();
            btnLogout.setToolTipText("????ng xu???t (" + tenNhanVien + " - " + DangNhap.nhanVienLogin.getMaNhanVien() + ")");
            btnLogout.addMouseListener(new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent me) {
                    logout();
                }
            });
            header.addItem(btnLogout, false);

            // btn Xem t??i kho???n ????ng nh???p
            NavBarButton btnSettingUser = new NavBarButton(new Rectangle(menuW - btnWidth, 0, btnWidth, headerH), "", "services_50px.png");
            btnSettingUser.setIconLocation(new Rectangle((btnWidth - iconSize) / 2, (headerH - iconSize) / 2, iconSize, iconSize));
            btnSettingUser.setBgDefault(new Color(headerBg, headerBg, headerBg));
            btnSettingUser.setBgHover(new Color(49, 49, 49));
            btnSettingUser.setToolTipText("T??i kho???n");
            btnSettingUser.addMouseListener(new MouseAdapter() {
                @Override
                public void mousePressed(MouseEvent me) {
                    new DoiMatKhauForm(DangNhap.taiKhoanLogin.getMaNhanVien()).setVisible(true);
                }
            });
            header.addItem(btnSettingUser, false);
        }

        // ========= Draggable ===========
        header.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent me) {
                px = me.getX();
                py = me.getY();
            }
        });

        header.addMouseMotionListener(new MouseMotionAdapter() {
            @Override
            public void mouseDragged(MouseEvent me) {
                setLocation(getLocation().x + me.getX() - px, getLocation().y + me.getY() - py);
            }
        });

        plContent.setLayout(new BorderLayout());
        JScrollPane workSpace = new JScrollPane(plContent);
        plContent.add(new BeginForm("Ch??o " + DangNhap.getTenNV()), BorderLayout.CENTER);
        add(header, BorderLayout.NORTH);
        add(scrollMenu, BorderLayout.WEST);
        add(workSpace, BorderLayout.CENTER);
        All.addWindowListener(new java.awt.event.WindowAdapter() {
            @Override
            public void windowClosed(java.awt.event.WindowEvent windowevent) {
                plContent.removeAll();
                plContent.add(new BeginForm("Ch??o " + DangNhap.getTenNV()), BorderLayout.CENTER);
                currentTab.setActive(false);
                revalidate();//refresh ui and layout
                repaint();
            }
        });
    }

    public static void BacktoMain() {
        All.setVisible(true);
        All.dispose();
    }

    public void doAction(String nameAction) throws DocumentException, IOException {
        plContent.removeAll();
        switch (nameAction) {
            case "Nh??n vi??n":
                if (nhanvien == null) {
                    nhanvien = new QuanliNhanVienForm();
                }
                plContent.setBackground(Color.DARK_GRAY);
                plContent.add(nhanvien, BorderLayout.CENTER);
                break;
            case "Kh??ch h??ng":
                if (khachhang == null) {
                    khachhang = new QuanLiKhachHangForm();
                }
                plContent.setBackground(Color.DARK_GRAY);
                plContent.add(khachhang, BorderLayout.CENTER);
                break;
            case "B???o h??nh":
                if (baohanh == null) {
                    baohanh = new QuanLiBaoHanh();
                }
                plContent.add(baohanh, BorderLayout.CENTER);
                break;
            case "Quy???n":
                if (quyen == null) {
                    quyen = new QuanLiQuyen();
                }
                plContent.add(quyen, BorderLayout.CENTER);
                break;
            case "Phi???u nh???p":
                if (phieunhap == null) {
                    phieunhap = new QuanLiPhieuNhap();
                }
                plContent.add(phieunhap, BorderLayout.CENTER);
                break;
            case "Lo???i s???n ph???m":
                if (lsanpham == null) {
                    lsanpham = new QuanLiLoaiSP();
                }
                plContent.add(lsanpham, BorderLayout.CENTER);
                break;
            case "Th????ng hi???u":
                if (thuonghieu == null) {
                    thuonghieu = new QuanLiThuongHieu();
                }
                plContent.add(thuonghieu, BorderLayout.CENTER);
                break;
            case "H??a ????n":
                if (hoadon == null) {
                    hoadon = new QuanLiHoaDon();
                }
                plContent.add(hoadon, BorderLayout.CENTER);
                break;
            case "S???n ph???m":
                if (sanpham == null) {
                    sanpham = new QuanLiSanPhamForm();

                }
                plContent.setBackground(Color.DARK_GRAY);
                plContent.add(sanpham, BorderLayout.CENTER);
                break;
            case "B??n h??ng":
                if (banhang == null) {
                    banhang = new BanHangForm();
                }
                plContent.setBackground(Color.DARK_GRAY);
                plContent.add(banhang, BorderLayout.CENTER);
                break;
            case "Khuy???n m??i":
                if (khuyenmai == null) {
                    khuyenmai = new QuanLiKhuyenMaiForm();
                }
                plContent.add(khuyenmai, BorderLayout.CENTER);
                break;
            case "Nh???p h??ng":
                ChonNhaCungCap.value = "";
                ChonNhaCungCap cNCC = new ChonNhaCungCap(null);
                cNCC.setVisible(true);
                cNCC.setAlwaysOnTop(true);
                cNCC.addWindowListener(new java.awt.event.WindowAdapter() {
                    @Override
                    public void windowClosed(java.awt.event.WindowEvent windowevent) {
                        nhaphang = new NhapHangForm();
                        //plContent.setBackground(Color.DARK_GRAY);
                        if (ChonNhaCungCap.value.equals("")) {
                            BacktoMain();
                        } else {
                            nhaphang.phieunhap.txNCC.setText(ChonNhaCungCap.value);
                            plContent.add(nhaphang, BorderLayout.CENTER);
                        }
                        status = 1;
                        revalidate();//refresh ui and layout
                        repaint();
                    }

                    @Override
                    public void windowOpened(java.awt.event.WindowEvent windowevent) {
                        status = 0;
                        //
                    }

                });

                break;
            case "Nh?? cung c???p": {
                if (nhacungcap == null) {
                    nhacungcap = new QuanLiNhaCungCap();
                }
                plContent.add(nhacungcap, BorderLayout.CENTER);
                break;
            }
            case "Th???ng k??":
                if (thongke == null) {
                    thongke = new ThongKe();
                }
                plContent.setBackground(Color.DARK_GRAY);
                plContent.add(thongke, BorderLayout.CENTER);
                break;
            case "C??ng c???":
                emptypage.setLabelText("C??ng c??? ??ang b???o tr??");
                plContent.add(emptypage, BorderLayout.CENTER);
                break;

            case "C??i ?????t":
                emptypage.setLabelText("C??i ?????t ??ang b???o tr??");
                plContent.add(emptypage, BorderLayout.CENTER);
                break;
        }
        headerTitle.setLabel(nameAction.toUpperCase());
        // https://stackoverflow.com/questions/12989388/switching-panels-with-menubar
        revalidate();//refresh ui and layout
        repaint();
    }

    @Override
    public void mouseClicked(MouseEvent e) {
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void mousePressed(MouseEvent e) {
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void mouseReleased(MouseEvent e) {
        if (e.getSource() instanceof NavBarButton && status == 1) {

            NavBarButton btn = (NavBarButton) e.getSource();
            if (currentTab != null) {
                currentTab.setActive(false);
            }

            btn.setActive(true);
            currentTab = btn;
            try {
                doAction(btn.text);
            } catch (DocumentException ex) {
                Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(Main.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void mouseEntered(MouseEvent e) {
        //throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void mouseExited(MouseEvent e) {
        // throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    private void logout() {
        int reply = JOptionPane.showConfirmDialog(getRootPane(),
                "B???n c?? ch???c mu???n ????ng xu???t kh???i " + DangNhap.getTenNV() + "?", "Ch?? ??",
                JOptionPane.YES_NO_OPTION);

        if (reply == JOptionPane.YES_OPTION) {
            //new WorkWithFile(LoginForm.saveFileName).write(""); // x??a d??? li???u ????ng nh???p
            new DangNhap().setVisible(true);
            this.dispose();
        }
    }
}
