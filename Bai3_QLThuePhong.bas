Attribute VB_Name = "Bai3_QLThuePhong"
' ============================================================
' BÀI 3: QUẢN LÝ CHO THUÊ PHÒNG - QLThuePhong.accdb
' Cách dùng: Alt+F11 > File > Import File > chọn file .bas này
'             Run > TaoTatCa_ThuePhong
' ============================================================
Sub TaoTatCa_ThuePhong()
    TaoBang_ThuePhong
    NhapDuLieu_ThuePhong
    MsgBox "Bai 3 HOAN THANH! Da tao bang va du lieu mau.", vbInformation
End Sub

Sub TaoBang_ThuePhong()
    Dim db As Database
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DROP TABLE THUE_PHONG"
    db.Execute "DROP TABLE PHONG"
    db.Execute "DROP TABLE KHACH_HANG"
    On Error GoTo 0

    db.Execute "CREATE TABLE PHONG (" & _
        "MaPhong TEXT(10) CONSTRAINT PK_Phong PRIMARY KEY, " & _
        "SoGiuong INTEGER, " & _
        "HoTenNVPhuTrach TEXT(50), " & _
        "GiaTien CURRENCY, " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE KHACH_HANG (" & _
        "MaKhachHang AUTOINCREMENT CONSTRAINT PK_KH PRIMARY KEY, " & _
        "TenKhachHang TEXT(50), " & _
        "DiaChi TEXT(100), " & _
        "SoDienThoai TEXT(12), " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE THUE_PHONG (" & _
        "MaKhachHang LONG, " & _
        "MaPhong TEXT(10), " & _
        "NgayLayPhong DATETIME, " & _
        "NgayTraPhong DATETIME, " & _
        "SoTienDaTra CURRENCY, " & _
        "GhiChu MEMO, " & _
        "CONSTRAINT PK_ThuePhong PRIMARY KEY (MaKhachHang, MaPhong, NgayLayPhong))"

    Dim rel As Relation, fld As Field
    On Error Resume Next
    Set rel = db.CreateRelation("FK_ThuePhong_KH", "KHACH_HANG", "THUE_PHONG", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaKhachHang")
    fld.ForeignName = "MaKhachHang"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_ThuePhong_Phong", "PHONG", "THUE_PHONG", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaPhong")
    fld.ForeignName = "MaPhong"
    rel.Fields.Append fld
    db.Relations.Append rel
    On Error GoTo 0
    Set db = Nothing
End Sub

Sub NhapDuLieu_ThuePhong()
    Dim db As Database
    Set db = CurrentDb
    db.Execute "INSERT INTO PHONG VALUES ('P101',1,'Nguyen Thi Lan',500000,'Phong don view ho boi')"
    db.Execute "INSERT INTO PHONG VALUES ('P201',2,'Tran Van Hung',800000,'Phong doi tang 2')"
    db.Execute "INSERT INTO PHONG VALUES ('P301',3,'Le Thi Mai',1200000,'Phong gia dinh tang 3')"

    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai) VALUES ('Nguyen Minh Giang','34 Hai Ba Trung Q1','0967890123')"
    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai) VALUES ('Tran Quoc Huy','67 Ly Thuong Kiet Q10','0978901234')"
    db.Execute "INSERT INTO KHACH_HANG (TenKhachHang,DiaChi,SoDienThoai) VALUES ('Dinh Thi Iris','90 Nguyen Trai Q5','0989012345')"

    db.Execute "INSERT INTO THUE_PHONG VALUES (1,'P101',#3/10/2025 14:00:00#,#3/12/2025 12:00:00#,1000000,'')"
    db.Execute "INSERT INTO THUE_PHONG VALUES (2,'P201',#3/11/2025 14:00:00#,#3/14/2025 12:00:00#,2400000,'')"
    db.Execute "INSERT INTO THUE_PHONG VALUES (3,'P301',#3/12/2025 14:00:00#,#3/15/2025 12:00:00#,3600000,'')"
    Set db = Nothing
End Sub
