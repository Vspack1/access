Attribute VB_Name = "Bai8_QLChuyenBay"
' ============================================================
' BÀI 8: QUẢN LÝ CHUYẾN BAY - QLChuyenBay.accdb
' Cách dùng: Alt+F11 > File > Import File > chọn file .bas này
'             Run > TaoTatCa_ChuyenBay
' ============================================================
Sub TaoTatCa_ChuyenBay()
    TaoBang_ChuyenBay
    NhapDuLieu_ChuyenBay
    MsgBox "Bai 8 HOAN THANH! Da tao bang va du lieu mau.", vbInformation
End Sub

Sub TaoBang_ChuyenBay()
    Dim db As Database
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DROP TABLE HK_CHUYEN_BAY"
    db.Execute "DROP TABLE TIEP_VIEN_CB"
    db.Execute "DROP TABLE PHI_CONG_PHU"
    db.Execute "DROP TABLE CHUYEN_BAY"
    db.Execute "DROP TABLE HANH_KHACH"
    db.Execute "DROP TABLE NHAN_VIEN_HK"
    db.Execute "DROP TABLE TUYEN_BAY"
    db.Execute "DROP TABLE MAY_BAY"
    On Error GoTo 0

    db.Execute "CREATE TABLE MAY_BAY (" & _
        "SoHieuMayBay TEXT(10) CONSTRAINT PK_MayBay PRIMARY KEY, " & _
        "HangSanXuat TEXT(50), " & _
        "NamSanXuat INTEGER, " & _
        "SoHieuModel TEXT(30), " & _
        "SoChoNgoi INTEGER, " & _
        "TrongTai DOUBLE)"

    db.Execute "CREATE TABLE TUYEN_BAY (" & _
        "MaTuyenBay TEXT(10) CONSTRAINT PK_TuyenBay PRIMARY KEY, " & _
        "DiemDi TEXT(50), " & _
        "DiemDen TEXT(50), " & _
        "KhoangCach DOUBLE)"

    db.Execute "CREATE TABLE NHAN_VIEN_HK (" & _
        "MaNVHK TEXT(10) CONSTRAINT PK_NVHK PRIMARY KEY, " & _
        "HoTen TEXT(100), " & _
        "NgaySinh DATETIME, " & _
        "DienThoai TEXT(20), " & _
        "ChucVu TEXT(20))"

    db.Execute "CREATE TABLE HANH_KHACH (" & _
        "MaHK TEXT(10) CONSTRAINT PK_HK PRIMARY KEY, " & _
        "HoTen TEXT(100), " & _
        "NgaySinh DATETIME, " & _
        "DienThoai TEXT(20), " & _
        "SoCMND TEXT(20))"

    db.Execute "CREATE TABLE CHUYEN_BAY (" & _
        "MaChuyenBay TEXT(10) CONSTRAINT PK_CB PRIMARY KEY, " & _
        "MaTuyenBay TEXT(10), " & _
        "SoHieuMayBay TEXT(10), " & _
        "NgayGioCatCanh DATETIME, " & _
        "NgayGioHaCanh DATETIME, " & _
        "MaPhiCongChinh TEXT(10))"

    db.Execute "CREATE TABLE PHI_CONG_PHU (" & _
        "MaChuyenBay TEXT(10), " & _
        "MaNVHK TEXT(10), " & _
        "CONSTRAINT PK_PCP PRIMARY KEY (MaChuyenBay, MaNVHK))"

    db.Execute "CREATE TABLE TIEP_VIEN_CB (" & _
        "MaChuyenBay TEXT(10), " & _
        "MaNVHK TEXT(10), " & _
        "CONSTRAINT PK_TVCB PRIMARY KEY (MaChuyenBay, MaNVHK))"

    db.Execute "CREATE TABLE HK_CHUYEN_BAY (" & _
        "MaChuyenBay TEXT(10), " & _
        "MaHK TEXT(10), " & _
        "SoGhe TEXT(5), " & _
        "HangVe TEXT(20), " & _
        "CONSTRAINT PK_HKCB PRIMARY KEY (MaChuyenBay, MaHK))"

    Dim rel As Relation, fld As Field
    On Error Resume Next
    Set rel = db.CreateRelation("FK_CB_Tuyen", "TUYEN_BAY", "CHUYEN_BAY", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaTuyenBay")
    fld.ForeignName = "MaTuyenBay"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_CB_MayBay", "MAY_BAY", "CHUYEN_BAY", dbRelationUpdateCascade)
    Set fld = rel.CreateField("SoHieuMayBay")
    fld.ForeignName = "SoHieuMayBay"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_CB_PhiCong", "NHAN_VIEN_HK", "CHUYEN_BAY", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaNVHK")
    fld.ForeignName = "MaPhiCongChinh"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_PCP_CB", "CHUYEN_BAY", "PHI_CONG_PHU", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaChuyenBay")
    fld.ForeignName = "MaChuyenBay"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_PCP_NV", "NHAN_VIEN_HK", "PHI_CONG_PHU", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaNVHK")
    fld.ForeignName = "MaNVHK"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_TVCB_CB", "CHUYEN_BAY", "TIEP_VIEN_CB", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaChuyenBay")
    fld.ForeignName = "MaChuyenBay"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_TVCB_NV", "NHAN_VIEN_HK", "TIEP_VIEN_CB", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaNVHK")
    fld.ForeignName = "MaNVHK"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_HKCB_CB", "CHUYEN_BAY", "HK_CHUYEN_BAY", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaChuyenBay")
    fld.ForeignName = "MaChuyenBay"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_HKCB_HK", "HANH_KHACH", "HK_CHUYEN_BAY", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaHK")
    fld.ForeignName = "MaHK"
    rel.Fields.Append fld
    db.Relations.Append rel
    On Error GoTo 0
    Set db = Nothing
End Sub

Sub NhapDuLieu_ChuyenBay()
    Dim db As Database
    Set db = CurrentDb
    db.Execute "INSERT INTO MAY_BAY VALUES ('VN-A321','Airbus',2018,'A321-200',180,73500)"
    db.Execute "INSERT INTO MAY_BAY VALUES ('VN-B738','Boeing',2015,'737-800',162,65000)"
    db.Execute "INSERT INTO MAY_BAY VALUES ('VN-A350','Airbus',2020,'A350-900',305,268000)"

    db.Execute "INSERT INTO TUYEN_BAY VALUES ('TB001','Ha Noi','TP.HCM',1137)"
    db.Execute "INSERT INTO TUYEN_BAY VALUES ('TB002','TP.HCM','Ha Noi',1137)"
    db.Execute "INSERT INTO TUYEN_BAY VALUES ('TB003','TP.HCM','Da Nang',610)"

    db.Execute "INSERT INTO NHAN_VIEN_HK VALUES ('PC001','Nguyen Thanh Long',#3/12/1975#,'0901111001','Phi cong')"
    db.Execute "INSERT INTO NHAN_VIEN_HK VALUES ('PC002','Tran Minh Quan',#7/22/1980#,'0901111002','Phi cong')"
    db.Execute "INSERT INTO NHAN_VIEN_HK VALUES ('PC003','Le Van Dung',#11/5/1982#,'0901111003','Phi cong')"
    db.Execute "INSERT INTO NHAN_VIEN_HK VALUES ('TV001','Pham Thi Huong',#4/18/1995#,'0901111004','Tiep vien')"
    db.Execute "INSERT INTO NHAN_VIEN_HK VALUES ('TV002','Hoang Thi Lan',#9/30/1997#,'0901111005','Tiep vien')"
    db.Execute "INSERT INTO NHAN_VIEN_HK VALUES ('TV003','Vo Thi Mai',#6/25/1996#,'0901111006','Tiep vien')"

    db.Execute "INSERT INTO HANH_KHACH VALUES ('HK001','Bui Van Tan',#2/14/1988#,'0911222001','079088001234')"
    db.Execute "INSERT INTO HANH_KHACH VALUES ('HK002','Ngo Thi Uyen',#8/20/1993#,'0911222002','079093005678')"
    db.Execute "INSERT INTO HANH_KHACH VALUES ('HK003','Dinh Minh Vu',#12/31/2000#,'0911222003','079000009012')"

    db.Execute "INSERT INTO CHUYEN_BAY VALUES ('CB001','TB001','VN-A321',#3/1/2025 6:00:00#,#3/1/2025 8:10:00#,'PC001')"
    db.Execute "INSERT INTO CHUYEN_BAY VALUES ('CB002','TB002','VN-B738',#3/1/2025 9:00:00#,#3/1/2025 11:05:00#,'PC002')"
    db.Execute "INSERT INTO CHUYEN_BAY VALUES ('CB003','TB003','VN-A350',#3/2/2025 7:30:00#,#3/2/2025 9:00:00#,'PC003')"

    db.Execute "INSERT INTO PHI_CONG_PHU VALUES ('CB001','PC002')"
    db.Execute "INSERT INTO PHI_CONG_PHU VALUES ('CB001','PC003')"
    db.Execute "INSERT INTO PHI_CONG_PHU VALUES ('CB002','PC001')"
    db.Execute "INSERT INTO PHI_CONG_PHU VALUES ('CB002','PC003')"

    db.Execute "INSERT INTO TIEP_VIEN_CB VALUES ('CB001','TV001')"
    db.Execute "INSERT INTO TIEP_VIEN_CB VALUES ('CB001','TV002')"
    db.Execute "INSERT INTO TIEP_VIEN_CB VALUES ('CB002','TV002')"
    db.Execute "INSERT INTO TIEP_VIEN_CB VALUES ('CB002','TV003')"
    db.Execute "INSERT INTO TIEP_VIEN_CB VALUES ('CB003','TV001')"
    db.Execute "INSERT INTO TIEP_VIEN_CB VALUES ('CB003','TV003')"

    db.Execute "INSERT INTO HK_CHUYEN_BAY VALUES ('CB001','HK001','12A','Pho thong')"
    db.Execute "INSERT INTO HK_CHUYEN_BAY VALUES ('CB001','HK002','12B','Pho thong')"
    db.Execute "INSERT INTO HK_CHUYEN_BAY VALUES ('CB002','HK003','1A','Thuong gia')"
    db.Execute "INSERT INTO HK_CHUYEN_BAY VALUES ('CB003','HK001','5C','Pho thong')"
    Set db = Nothing
End Sub
