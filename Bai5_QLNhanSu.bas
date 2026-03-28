Attribute VB_Name = "Bai5_QLNhanSu"
' ============================================================
' BÀI 5: QUẢN LÝ NHÂN SỰ DỰ ÁN - QLNhanSu.accdb
' Cách dùng: Alt+F11 > File > Import File > chọn file .bas này
'             Run > TaoTatCa_NhanSu
' ============================================================
Sub TaoTatCa_NhanSu()
    TaoBang_NhanSu
    NhapDuLieu_NhanSu
    MsgBox "Bai 5 HOAN THANH! Da tao bang va du lieu mau.", vbInformation
End Sub

Sub TaoBang_NhanSu()
    Dim db As Database
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DROP TABLE THAM_GIA"
    db.Execute "DROP TABLE NHAN_VIEN"
    db.Execute "DROP TABLE DU_AN"
    On Error GoTo 0

    db.Execute "CREATE TABLE NHAN_VIEN (" & _
        "MaNhanVien TEXT(10) CONSTRAINT PK_NV PRIMARY KEY, " & _
        "TenNhanVien TEXT(50), " & _
        "BangCap TEXT(50), " & _
        "NamSinh INTEGER, " & _
        "DiaChi TEXT(100), " & _
        "ChucVu TEXT(50))"

    db.Execute "CREATE TABLE DU_AN (" & _
        "MaDuAn TEXT(10) CONSTRAINT PK_DA PRIMARY KEY, " & _
        "TenDuAn TEXT(100), " & _
        "NgayDuKienBatDau DATETIME, " & _
        "NgayBatDau DATETIME, " & _
        "NgayDuKienKetThuc DATETIME, " & _
        "NgayKetThuc DATETIME, " & _
        "GhiChu MEMO)"

    db.Execute "CREATE TABLE THAM_GIA (" & _
        "MaNhanVien TEXT(10), " & _
        "MaDuAn TEXT(10), " & _
        "NgayBatDauThamGia DATETIME, " & _
        "NgayKetThucThamGia DATETIME, " & _
        "NhiemVu TEXT(100), " & _
        "DanhGiaKetQua TEXT(50), " & _
        "CONSTRAINT PK_ThamGia PRIMARY KEY (MaNhanVien, MaDuAn, NgayBatDauThamGia))"

    Dim rel As Relation, fld As Field
    On Error Resume Next
    Set rel = db.CreateRelation("FK_ThamGia_NV", "NHAN_VIEN", "THAM_GIA", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaNhanVien")
    fld.ForeignName = "MaNhanVien"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_ThamGia_DA", "DU_AN", "THAM_GIA", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaDuAn")
    fld.ForeignName = "MaDuAn"
    rel.Fields.Append fld
    db.Relations.Append rel
    On Error GoTo 0
    Set db = Nothing
End Sub

Sub NhapDuLieu_NhanSu()
    Dim db As Database
    Set db = CurrentDb
    db.Execute "INSERT INTO NHAN_VIEN VALUES ('NV001','Nguyen Van Nam','Ky su CNTT',1990,'12 Le Van Viet Q9','Lap trinh vien')"
    db.Execute "INSERT INTO NHAN_VIEN VALUES ('NV002','Tran Thi Oanh','Thac si CNTT',1988,'34 Quang Trung GV','Truong nhom')"
    db.Execute "INSERT INTO NHAN_VIEN VALUES ('NV003','Le Quang Phuc','Cu nhan CNTT',1995,'56 To Hien Thanh Q10','Kiem thu')"

    db.Execute "INSERT INTO DU_AN VALUES ('DA001','He thong quan ly ban hang',#1/1/2025#,#1/15/2025#,#6/30/2025#,Null,'Du an cho khach hang A')"
    db.Execute "INSERT INTO DU_AN VALUES ('DA002','App dat do an online',#2/1/2025#,#2/10/2025#,#8/31/2025#,Null,'Startup noi bo')"
    db.Execute "INSERT INTO DU_AN VALUES ('DA003','Website thuong mai dien tu',#3/1/2025#,#3/5/2025#,#9/30/2025#,Null,'Du an cho khach hang B')"

    db.Execute "INSERT INTO THAM_GIA VALUES ('NV001','DA001',#1/15/2025#,#4/15/2025#,'Backend developer','Tot')"
    db.Execute "INSERT INTO THAM_GIA VALUES ('NV002','DA001',#1/15/2025#,Null,'Truong nhom',Null)"
    db.Execute "INSERT INTO THAM_GIA VALUES ('NV003','DA002',#2/10/2025#,Null,'Tester',Null)"
    db.Execute "INSERT INTO THAM_GIA VALUES ('NV001','DA002',#2/15/2025#,#3/1/2025#,'Frontend developer','Kha')"
    ' NV001 tham gia lai DA002
    db.Execute "INSERT INTO THAM_GIA VALUES ('NV001','DA002',#3/10/2025#,Null,'Frontend developer',Null)"
    Set db = Nothing
End Sub
