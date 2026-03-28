Attribute VB_Name = "Bai6_QLThuVien"
' ============================================================
' BÀI 6: QUẢN LÝ THƯ VIỆN - QLThuVien.accdb
' Cách dùng: Alt+F11 > File > Import File > chọn file .bas này
'             Run > TaoTatCa_ThuVien
' ============================================================
Sub TaoTatCa_ThuVien()
    TaoBang_ThuVien
    NhapDuLieu_ThuVien
    MsgBox "Bai 6 HOAN THANH! Da tao bang va du lieu mau.", vbInformation
End Sub

Sub TaoBang_ThuVien()
    Dim db As Database
    Set db = CurrentDb
    On Error Resume Next
    db.Execute "DROP TABLE MUON_TRA"
    db.Execute "DROP TABLE SACH_TAC_GIA"
    db.Execute "DROP TABLE SACH"
    db.Execute "DROP TABLE TAC_GIA"
    db.Execute "DROP TABLE NHA_XUAT_BAN"
    db.Execute "DROP TABLE DOC_GIA"
    On Error GoTo 0

    db.Execute "CREATE TABLE NHA_XUAT_BAN (" & _
        "MaNXB TEXT(10) CONSTRAINT PK_NXB PRIMARY KEY, " & _
        "TenNXB TEXT(100), " & _
        "DiaChi TEXT(200), " & _
        "DienThoai TEXT(20))"

    db.Execute "CREATE TABLE TAC_GIA (" & _
        "MaTacGia TEXT(10) CONSTRAINT PK_TacGia PRIMARY KEY, " & _
        "TenTacGia TEXT(100), " & _
        "NgaySinh DATETIME, " & _
        "QuocTich TEXT(50))"

    db.Execute "CREATE TABLE SACH (" & _
        "MaSach TEXT(10) CONSTRAINT PK_Sach PRIMARY KEY, " & _
        "TenSach TEXT(100), " & _
        "MaNXB TEXT(10), " & _
        "NamXuatBan INTEGER)"

    db.Execute "CREATE TABLE SACH_TAC_GIA (" & _
        "MaSach TEXT(10), " & _
        "MaTacGia TEXT(10), " & _
        "CONSTRAINT PK_STG PRIMARY KEY (MaSach, MaTacGia))"

    db.Execute "CREATE TABLE DOC_GIA (" & _
        "MaDocGia AUTOINCREMENT CONSTRAINT PK_DocGia PRIMARY KEY, " & _
        "TenDocGia TEXT(100), " & _
        "NgaySinh DATETIME, " & _
        "DiaChi TEXT(200), " & _
        "DienThoai TEXT(20), " & _
        "Email TEXT(100))"

    db.Execute "CREATE TABLE MUON_TRA (" & _
        "MaDocGia LONG, " & _
        "MaSach TEXT(10), " & _
        "NgayGioMuon DATETIME, " & _
        "NgayGioTra DATETIME, " & _
        "TinhTrangKhiMuon TEXT(50), " & _
        "TinhTrangKhiTra TEXT(50), " & _
        "CONSTRAINT PK_MuonTra PRIMARY KEY (MaDocGia, MaSach, NgayGioMuon))"

    Dim rel As Relation, fld As Field
    On Error Resume Next
    Set rel = db.CreateRelation("FK_Sach_NXB", "NHA_XUAT_BAN", "SACH", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaNXB")
    fld.ForeignName = "MaNXB"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_STG_Sach", "SACH", "SACH_TAC_GIA", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaSach")
    fld.ForeignName = "MaSach"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_STG_TacGia", "TAC_GIA", "SACH_TAC_GIA", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaTacGia")
    fld.ForeignName = "MaTacGia"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_MT_DocGia", "DOC_GIA", "MUON_TRA", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaDocGia")
    fld.ForeignName = "MaDocGia"
    rel.Fields.Append fld
    db.Relations.Append rel

    Set rel = db.CreateRelation("FK_MT_Sach", "SACH", "MUON_TRA", dbRelationUpdateCascade)
    Set fld = rel.CreateField("MaSach")
    fld.ForeignName = "MaSach"
    rel.Fields.Append fld
    db.Relations.Append rel
    On Error GoTo 0
    Set db = Nothing
End Sub

Sub NhapDuLieu_ThuVien()
    Dim db As Database
    Set db = CurrentDb
    db.Execute "INSERT INTO NHA_XUAT_BAN VALUES ('NXB001','NXB Kim Dong','248A No Trang Long BT','02838412511')"
    db.Execute "INSERT INTO NHA_XUAT_BAN VALUES ('NXB002','NXB Giao Duc','81 Tran Hung Dao Q1','02838222361')"
    db.Execute "INSERT INTO NHA_XUAT_BAN VALUES ('NXB003','NXB Tre','161B Ly Chinh Thang Q3','02839316289')"

    db.Execute "INSERT INTO TAC_GIA VALUES ('TG001','Nguyen Nhat Anh',#5/7/1955#,'Viet Nam')"
    db.Execute "INSERT INTO TAC_GIA VALUES ('TG002','Nam Cao',#10/29/1917#,'Viet Nam')"
    db.Execute "INSERT INTO TAC_GIA VALUES ('TG003','To Hoai',#9/27/1920#,'Viet Nam')"

    db.Execute "INSERT INTO SACH VALUES ('ST001','Toi Thay Hoa Vang Tren Co Xanh','NXB003',2010)"
    db.Execute "INSERT INTO SACH VALUES ('ST002','Chi Pheo','NXB002',1998)"
    db.Execute "INSERT INTO SACH VALUES ('ST003','De Men Phieu Luu Ky','NXB001',2005)"

    db.Execute "INSERT INTO SACH_TAC_GIA VALUES ('ST001','TG001')"
    db.Execute "INSERT INTO SACH_TAC_GIA VALUES ('ST002','TG002')"
    db.Execute "INSERT INTO SACH_TAC_GIA VALUES ('ST003','TG003')"

    db.Execute "INSERT INTO DOC_GIA (TenDocGia,NgaySinh,DiaChi,DienThoai,Email) VALUES ('Tran Thi Quynh',#3/15/2000#,'45 Phan Dinh Phung QPN','0911000111','quynh@gmail.com')"
    db.Execute "INSERT INTO DOC_GIA (TenDocGia,NgaySinh,DiaChi,DienThoai,Email) VALUES ('Nguyen Van Rong',#7/20/1998#,'67 Bui Thi Xuan Q1','0911000222','rong@gmail.com')"
    db.Execute "INSERT INTO DOC_GIA (TenDocGia,NgaySinh,DiaChi,DienThoai,Email) VALUES ('Vo Thi Suong',#11/5/2001#,'89 Dinh Tien Hoang QBT','0911000333','suong@gmail.com')"

    db.Execute "INSERT INTO MUON_TRA VALUES (1,'ST001',#3/1/2025 8:00:00#,#3/15/2025 8:00:00#,'Tot','Tot')"
    db.Execute "INSERT INTO MUON_TRA VALUES (2,'ST002',#3/5/2025 9:00:00#,Null,'Tot',Null)"
    db.Execute "INSERT INTO MUON_TRA VALUES (3,'ST003',#3/10/2025 10:00:00#,Null,'Binh thuong',Null)"
    db.Execute "INSERT INTO MUON_TRA VALUES (1,'ST001',#3/20/2025 8:00:00#,Null,'Tot',Null)"
    Set db = Nothing
End Sub
