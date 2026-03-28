Attribute VB_Name = "Bai2_QLThueDia"
' ============================================================
' BÀI 2: QUẢN LÝ CHO THUÊ ĐĨA - QLThueDia.accdb
' Cách dùng: Alt+F11 > File > Import File > chọn file .bas này
'             Run > TaoTatCa_ThueDia
' ============================================================
Sub TaoTatCa_ThueDia()
    TaoBang_ThueDia
    NhapDuLieu_ThueDia
    MsgBox "Bai 2 HOAN THANH! Da tao bang va du lieu mau.", vbInformation
End Sub

Sub TaoBang_ThueDia()
    Dim db As Database
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DROP TABLE THUE_DIA"
    db.Execute "DROP TABLE DIA"
    db.Execute "DROP TABLE KHACH_HANG"
    On Error GoTo 0

    db.Execute "CREATE TABLE DIA (" & _
        "MaDia TEXT(10) CONSTRAINT PK_Dia PRIMARY KEY, " & _
        "TenDia TEXT(50), " & _
        "TheLoai TEXT(20), " & _
        "TenNuocSanXuat TEXT(20), " & _
        "GiaMuaVao CURRENCY, " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE KHACH_HANG (" & _
        "MaKhachHang AUTOINCREMENT CONSTRAINT PK_KH PRIMARY KEY, " & _
        "TenKhachHang TEXT(50), " & _
        "DiaChi TEXT(100), " & _
        "SoDienThoai TEXT(12), " & _
        "TheLoaiYeuThich TEXT(20), " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE THUE_DIA (" & _
        "MaKhachHang LONG, " & _
        "MaDia TEXT(10), " & _
        "NgayGioThue DATETIME, " & _
        "NgayGioTra DATETIME, " & _
        "SoTienThu CURRENCY, " & _
        "GhiChu MEMO, " & _
        "CONSTRAINT PK_ThueDia PRIMARY KEY (MaKhachHang, MaDia, NgayGioThue))"

    Dim rel As Relation, fld As Field
    On Error Resume Next
    Set rel = db.CreateRelation("FK_ThueDia_KH", "KHACH_HANG", "THUE_DIA", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaKhachHang")
    fld.ForeignName = "MaKhachHang"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_ThueDia_Dia", "DIA", "THUE_DIA", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaDia")
    fld.ForeignName = "MaDia"
    rel.Fields.Append fld
    db.Relations.Append rel
    On Error GoTo 0
    Set db = Nothing
End Sub

Sub NhapDuLieu_ThueDia()
    Dim db As Database
    Set db = CurrentDb
    db.Execute "INSERT INTO DIA VALUES ('D001','Avengers Endgame','Phim','My',120000,'')"
    db.Execute "INSERT INTO DIA VALUES ('D002','Son Tung MTP Album','Nhac','Viet Nam',80000,'')"
    db.Execute "INSERT INTO DIA VALUES ('D003','Your Name','Phim','Nhat Ban',100000,'')"

    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai,TheLoaiYeuThich) VALUES ('Pham Thi Dung','12 Dien Bien Phu Q3','0934567890','Phim')"
    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai,TheLoaiYeuThich) VALUES ('Hoang Van Em','56 CMT8 Q10','0945678901','Nhac')"
    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai,TheLoaiYeuThich) VALUES ('Vo Thi Phuong','89 Phan Xich Long BT','0956789012','Phim')"

    db.Execute "INSERT INTO THUE_DIA VALUES (1,'D001',#3/1/2025 10:00:00#,#3/3/2025 10:00:00#,20000,'')"
    db.Execute "INSERT INTO THUE_DIA VALUES (2,'D002',#3/2/2025 11:00:00#,#3/4/2025 11:00:00#,15000,'')"
    db.Execute "INSERT INTO THUE_DIA VALUES (3,'D003',#3/3/2025 14:00:00#,#3/6/2025 14:00:00#,20000,'')"
    Set db = Nothing
End Sub
