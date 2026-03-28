Attribute VB_Name = "Bai7_QLBanHang"
' ============================================================
' BÀI 7: QUẢN LÝ BÁN HÀNG - QLBanHang.accdb
' Cách dùng: Alt+F11 > File > Import File > chọn file .bas này
'             Run > TaoTatCa_BanHang
' ============================================================
Sub TaoTatCa_BanHang()
    TaoBang_BanHang
    NhapDuLieu_BanHang
    MsgBox "Bai 7 HOAN THANH! Da tao bang va du lieu mau.", vbInformation
End Sub

Sub TaoBang_BanHang()
    Dim db As Database
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DROP TABLE CHI_TIET_DH"
    db.Execute "DROP TABLE DON_HANG"
    db.Execute "DROP TABLE HANG_HOA"
    db.Execute "DROP TABLE KHACH_HANG"
    db.Execute "DROP TABLE NHAN_VIEN"
    On Error GoTo 0

    db.Execute "CREATE TABLE KHACH_HANG (" & _
        "MaKH TEXT(10) CONSTRAINT PK_KH PRIMARY KEY, " & _
        "Ho TEXT(30), " & _
        "Ten TEXT(20), " & _
        "NgaySinh DATETIME, " & _
        "GioiTinh TEXT(3), " & _
        "DiaChi TEXT(100), " & _
        "SoDienThoai TEXT(12))"

    db.Execute "CREATE TABLE NHAN_VIEN (" & _
        "MaNV TEXT(10) CONSTRAINT PK_NV PRIMARY KEY, " & _
        "Ho TEXT(30), " & _
        "Ten TEXT(20), " & _
        "NgaySinh DATETIME, " & _
        "GioiTinh TEXT(3), " & _
        "DiaChi TEXT(100), " & _
        "SoDienThoai TEXT(12))"

    db.Execute "CREATE TABLE HANG_HOA (" & _
        "MaHH TEXT(10) CONSTRAINT PK_HH PRIMARY KEY, " & _
        "TenHH TEXT(100), " & _
        "DonViTinh TEXT(20), " & _
        "DonGiaNiemYet CURRENCY)"

    db.Execute "CREATE TABLE DON_HANG (" & _
        "MaDonHang TEXT(10) CONSTRAINT PK_DH PRIMARY KEY, " & _
        "NgayMua DATETIME, " & _
        "TienVanChuyen CURRENCY, " & _
        "MaKH TEXT(10), " & _
        "MaNV TEXT(10))"

    db.Execute "CREATE TABLE CHI_TIET_DH (" & _
        "MaDonHang TEXT(10), " & _
        "MaHH TEXT(10), " & _
        "DonGiaBan CURRENCY, " & _
        "SoLuong INTEGER, " & _
        "CONSTRAINT PK_CTDH PRIMARY KEY (MaDonHang, MaHH))"

    Dim rel As Relation, fld As Field
    On Error Resume Next
    Set rel = db.CreateRelation("FK_DH_KH", "KHACH_HANG", "DON_HANG", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaKH")
    fld.ForeignName = "MaKH"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_DH_NV", "NHAN_VIEN", "DON_HANG", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaNV")
    fld.ForeignName = "MaNV"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_CTDH_DH", "DON_HANG", "CHI_TIET_DH", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaDonHang")
    fld.ForeignName = "MaDonHang"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_CTDH_HH", "HANG_HOA", "CHI_TIET_DH", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaHH")
    fld.ForeignName = "MaHH"
    rel.Fields.Append fld
    db.Relations.Append rel
    On Error GoTo 0
    Set db = Nothing
End Sub

Sub NhapDuLieu_BanHang()
    Dim db As Database
    Set db = CurrentDb
    db.Execute "INSERT INTO KHACH_HANG VALUES ('KH001','Nguyen','Thang',#6/10/1985#,'Nam','10 Vo Van Tan Q3','0901001001')"
    db.Execute "INSERT INTO KHACH_HANG VALUES ('KH002','Tran','Uyen',#9/25/1992#,'Nu','20 Tran Quang Khai Q1','0902002002')"
    db.Execute "INSERT INTO KHACH_HANG VALUES ('KH003','Le','Vinh',#12/15/1990#,'Nam','30 Hung Vuong Q5','0903003003')"

    db.Execute "INSERT INTO NHAN_VIEN VALUES ('NV001','Pham','Xuan',#4/20/1988#,'Nam','15 Ngo Quyen Q10','0904004004')"
    db.Execute "INSERT INTO NHAN_VIEN VALUES ('NV002','Hoang','Yen',#8/30/1995#,'Nu','25 Le Thanh Ton Q1','0905005005')"
    db.Execute "INSERT INTO NHAN_VIEN VALUES ('NV003','Do','Zoey',#1/12/1993#,'Nu','35 Nam Ky Khoi Nghia Q3','0906006006')"

    db.Execute "INSERT INTO HANG_HOA VALUES ('HH001','Laptop Dell Inspiron','Cai',15000000)"
    db.Execute "INSERT INTO HANG_HOA VALUES ('HH002','Chuot khong day Logitech','Cai',350000)"
    db.Execute "INSERT INTO HANG_HOA VALUES ('HH003','Ban phim co Keychron','Cai',1800000)"
    db.Execute "INSERT INTO HANG_HOA VALUES ('HH004','Man hinh LG 24 inch','Cai',4500000)"

    db.Execute "INSERT INTO DON_HANG VALUES ('DH001',#3/1/2025#,50000,'KH001','NV001')"
    db.Execute "INSERT INTO DON_HANG VALUES ('DH002',#3/5/2025#,30000,'KH002','NV002')"
    db.Execute "INSERT INTO DON_HANG VALUES ('DH003',#3/10/2025#,0,'KH003','NV001')"

    db.Execute "INSERT INTO CHI_TIET_DH VALUES ('DH001','HH001',14500000,1)"
    db.Execute "INSERT INTO CHI_TIET_DH VALUES ('DH001','HH002',320000,2)"
    db.Execute "INSERT INTO CHI_TIET_DH VALUES ('DH002','HH003',1750000,1)"
    db.Execute "INSERT INTO CHI_TIET_DH VALUES ('DH002','HH004',4300000,1)"
    db.Execute "INSERT INTO CHI_TIET_DH VALUES ('DH003','HH002',320000,3)"
    Set db = Nothing
End Sub
