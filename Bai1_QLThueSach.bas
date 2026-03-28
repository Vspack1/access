Attribute VB_Name = "Bai1_QLThueSach"
' ============================================================
' BÀI 1: QUẢN LÝ CHO THUÊ SÁCH - QLThueSach.accdb
' Cách dùng: Alt+F11 > File > Import File > chọn file .bas này
'             Run > TaoTatCa_ThueSach
' ============================================================
Sub TaoTatCa_ThueSach()
    TaoBang_ThueSach
    NhapDuLieu_ThueSach
    MsgBox "Bai 1 HOAN THANH! Da tao bang va du lieu mau.", vbInformation
End Sub

Sub TaoBang_ThueSach()
    Dim db As Database
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DROP TABLE THUE_SACH"
    db.Execute "DROP TABLE SACH"
    db.Execute "DROP TABLE LOAI_SACH"
    db.Execute "DROP TABLE KHACH_HANG"
    On Error GoTo 0

    db.Execute "CREATE TABLE LOAI_SACH (" & _
        "MaLoaiSach TEXT(10) CONSTRAINT PK_LoaiSach PRIMARY KEY, " & _
        "TenLoaiSach TEXT(100), " & _
        "MieuTaLoaiSach TEXT(255))"

    db.Execute "CREATE TABLE SACH (" & _
        "MaSach TEXT(10) CONSTRAINT PK_Sach PRIMARY KEY, " & _
        "TenSach TEXT(50), " & _
        "TacGia TEXT(50), " & _
        "TenNXB TEXT(50), " & _
        "MaLoaiSach TEXT(10), " & _
        "GiaMuaVao CURRENCY, " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE KHACH_HANG (" & _
        "MaKhachHang AUTOINCREMENT CONSTRAINT PK_KH PRIMARY KEY, " & _
        "TenKhachHang TEXT(50), " & _
        "DiaChi TEXT(100), " & _
        "SoDienThoai TEXT(12), " & _
        "LoaiSachYeuThich TEXT(100), " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE THUE_SACH (" & _
        "MaKhachHang LONG, " & _
        "MaSach TEXT(10), " & _
        "NgayGioMuon DATETIME, " & _
        "NgayGioTra DATETIME, " & _
        "SoTienThue CURRENCY, " & _
        "GhiChu MEMO, " & _
        "CONSTRAINT PK_ThueSach PRIMARY KEY (MaKhachHang, MaSach, NgayGioMuon))"

    Dim rel As Relation, fld As Field
    On Error Resume Next
    Set rel = db.CreateRelation("FK_Sach_LoaiSach", "LOAI_SACH", "SACH", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaLoaiSach")
    fld.ForeignName = "MaLoaiSach"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_ThueSach_KH", "KHACH_HANG", "THUE_SACH", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaKhachHang")
    fld.ForeignName = "MaKhachHang"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_ThueSach_Sach", "SACH", "THUE_SACH", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaSach")
    fld.ForeignName = "MaSach"
    rel.Fields.Append fld
    db.Relations.Append rel
    On Error GoTo 0
    Set db = Nothing
End Sub

Sub NhapDuLieu_ThueSach()
    Dim db As Database
    Set db = CurrentDb
    db.Execute "INSERT INTO LOAI_SACH VALUES ('LS001','Van hoc','Cac tac pham van hoc')"
    db.Execute "INSERT INTO LOAI_SACH VALUES ('LS002','Khoa hoc','Sach khoa hoc tu nhien')"
    db.Execute "INSERT INTO LOAI_SACH VALUES ('LS003','Thieu nhi','Sach danh cho tre em')"

    db.Execute "INSERT INTO SACH VALUES ('S001','Doraemon Tap 1','Fujiko F. Fujio','NXB Kim Dong','LS003',35000,'')"
    db.Execute "INSERT INTO SACH VALUES ('S002','So Do','Vu Trong Phung','NXB Van hoc','LS001',45000,'Tieu thuyet')"
    db.Execute "INSERT INTO SACH VALUES ('S003','Vat Ly Dai Cuong','Luong Duyen Binh','NXB Giao duc','LS002',80000,'')"

    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai,LoaiSachYeuThich) VALUES ('Nguyen Van An','123 Le Loi Q1','0901234567','Van hoc')"
    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai,LoaiSachYeuThich) VALUES ('Tran Thi Binh','45 Nguyen Hue Q1','0912345678','Thieu nhi')"
    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai,LoaiSachYeuThich) VALUES ('Le Minh Cuong','78 Tran Hung Dao Q5','0923456789','Khoa hoc')"

    db.Execute "INSERT INTO THUE_SACH VALUES (1,'S002',#3/1/2025 8:00:00#,#3/8/2025 8:00:00#,10000,'')"
    db.Execute "INSERT INTO THUE_SACH VALUES (2,'S001',#3/2/2025 9:00:00#,#3/5/2025 9:00:00#,5000,'')"
    db.Execute "INSERT INTO THUE_SACH VALUES (3,'S003',#3/3/2025 10:00:00#,#3/10/2025 10:00:00#,15000,'')"
    Set db = Nothing
End Sub
