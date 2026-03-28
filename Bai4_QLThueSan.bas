Attribute VB_Name = "Bai4_QLThueSan"
' ============================================================
' BÀI 4: QUẢN LÝ CHO THUÊ SÂN QUẦN VỢT - QLThueSan.accdb
' Cách dùng: Alt+F11 > File > Import File > chọn file .bas này
'             Run > TaoTatCa_ThueSan
' ============================================================
Sub TaoTatCa_ThueSan()
    TaoBang_ThueSan
    NhapDuLieu_ThueSan
    MsgBox "Bai 4 HOAN THANH! Da tao bang va du lieu mau.", vbInformation
End Sub

Sub TaoBang_ThueSan()
    Dim db As Database
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DROP TABLE THUE_SAN"
    db.Execute "DROP TABLE SAN"
    db.Execute "DROP TABLE KHACH_HANG"
    On Error GoTo 0

    db.Execute "CREATE TABLE SAN (" & _
        "MaSan TEXT(10) CONSTRAINT PK_San PRIMARY KEY, " & _
        "TinhTrang TEXT(50), " & _
        "GiaThue CURRENCY, " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE KHACH_HANG (" & _
        "MaKhachHang AUTOINCREMENT CONSTRAINT PK_KH PRIMARY KEY, " & _
        "TenKhachHang TEXT(50), " & _
        "DiaChi TEXT(100), " & _
        "SoDienThoai TEXT(12), " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE THUE_SAN (" & _
        "MaKhachHang LONG, " & _
        "MaSan TEXT(10), " & _
        "NgayGioThue DATETIME, " & _
        "NgayGioTra DATETIME, " & _
        "SoTienThu CURRENCY, " & _
        "GhiChu MEMO, " & _
        "CONSTRAINT PK_ThueSan PRIMARY KEY (MaKhachHang, MaSan, NgayGioThue))"

    Dim rel As Relation, fld As Field
    On Error Resume Next
    Set rel = db.CreateRelation("FK_ThueSan_KH", "KHACH_HANG", "THUE_SAN", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaKhachHang")
    fld.ForeignName = "MaKhachHang"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_ThueSan_San", "SAN", "THUE_SAN", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaSan")
    fld.ForeignName = "MaSan"
    rel.Fields.Append fld
    db.Relations.Append rel
    On Error GoTo 0
    Set db = Nothing
End Sub

Sub NhapDuLieu_ThueSan()
    Dim db As Database
    Set db = CurrentDb
    db.Execute "INSERT INTO SAN VALUES ('SAN01','Tot',150000,'San trong nha co den')"
    db.Execute "INSERT INTO SAN VALUES ('SAN02','Binh thuong',100000,'San ngoai troi')"
    db.Execute "INSERT INTO SAN VALUES ('SAN03','Dang sua chua',0,'Tam ngung hoat dong')"

    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai) VALUES ('Ly Van Khoa','15 Dinh Tien Hoang BT','0901111222')"
    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai) VALUES ('Pham Minh Long','28 No Trang Long BT','0902222333')"
    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai) VALUES ('Dang Thi My','39 Phan Van Tri GV','0903333444')"

    db.Execute "INSERT INTO THUE_SAN VALUES (1,'SAN01',#3/15/2025 7:00:00#,#3/15/2025 9:00:00#,300000,'')"
    db.Execute "INSERT INTO THUE_SAN VALUES (2,'SAN02',#3/15/2025 8:00:00#,#3/15/2025 10:00:00#,200000,'')"
    db.Execute "INSERT INTO THUE_SAN VALUES (3,'SAN01',#3/16/2025 6:00:00#,#3/16/2025 8:00:00#,300000,'')"
    Set db = Nothing
End Sub
