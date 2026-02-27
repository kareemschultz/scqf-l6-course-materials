"""
UNIT B: F1FJ 12 - Access Database (.accdb)
v2: Simplified form creation with proper COM handling
"""
import win32com.client
import os
import time
import pythoncom
import gc

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FJ12_Spreadsheet_Database")
ACCDB_PATH = os.path.join(SRC, "F1FJ12_Database.accdb")
os.makedirs(SRC, exist_ok=True)

dbText = 10
dbLong = 4
dbCurrency = 5
dbDate = 8
dbInteger = 3
dbAutoIncrField = 16


def create_tables_and_data():
    """Phase 1: Create tables, relationships, data, queries using DAO"""
    pythoncom.CoInitialize()

    if os.path.exists(ACCDB_PATH):
        os.remove(ACCDB_PATH)

    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False

    try:
        db_engine = access.DBEngine
        ws = db_engine.Workspaces(0)
        db = ws.CreateDatabase(ACCDB_PATH, ";LANGID=0x0409;CP=1252;COUNTRY=0", 64)

        # --- Books table ---
        td = db.CreateTableDef("Books")
        fld = td.CreateField("BookID", dbLong); fld.Attributes = dbAutoIncrField; td.Fields.Append(fld)
        for n, t, s in [("Title",dbText,100),("Author",dbText,80),("Genre",dbText,50),("ISBN",dbText,20),("Price",dbCurrency,0),("YearPublished",dbInteger,0)]:
            f = td.CreateField(n, t)
            if t == dbText: f.Size = s
            td.Fields.Append(f)
        idx = td.CreateIndex("PrimaryKey"); idx.Primary = True; idx.Fields.Append(idx.CreateField("BookID")); td.Indexes.Append(idx)
        db.TableDefs.Append(td)

        # --- Members table ---
        td = db.CreateTableDef("Members")
        fld = td.CreateField("MemberID", dbLong); fld.Attributes = dbAutoIncrField; td.Fields.Append(fld)
        for n, t, s in [("FirstName",dbText,50),("LastName",dbText,50),("Email",dbText,100),("Phone",dbText,20),("MembershipDate",dbDate,0)]:
            f = td.CreateField(n, t)
            if t == dbText: f.Size = s
            td.Fields.Append(f)
        idx = td.CreateIndex("PrimaryKey"); idx.Primary = True; idx.Fields.Append(idx.CreateField("MemberID")); td.Indexes.Append(idx)
        db.TableDefs.Append(td)

        # --- Borrowing table ---
        td = db.CreateTableDef("Borrowing")
        fld = td.CreateField("BorrowID", dbLong); fld.Attributes = dbAutoIncrField; td.Fields.Append(fld)
        for n, t in [("MemberID",dbLong),("BookID",dbLong),("BorrowDate",dbDate),("ReturnDate",dbDate)]:
            td.Fields.Append(td.CreateField(n, t))
        idx = td.CreateIndex("PrimaryKey"); idx.Primary = True; idx.Fields.Append(idx.CreateField("BorrowID")); td.Indexes.Append(idx)
        db.TableDefs.Append(td)

        # --- Patients table ---
        td = db.CreateTableDef("Patients")
        fld = td.CreateField("PatientID", dbLong); fld.Attributes = dbAutoIncrField; td.Fields.Append(fld)
        for n, t, s in [("FirstName",dbText,50),("LastName",dbText,50),("DOB",dbDate,0),("Phone",dbText,20),("Address",dbText,150)]:
            f = td.CreateField(n, t)
            if t == dbText: f.Size = s
            td.Fields.Append(f)
        idx = td.CreateIndex("PrimaryKey"); idx.Primary = True; idx.Fields.Append(idx.CreateField("PatientID")); td.Indexes.Append(idx)
        db.TableDefs.Append(td)

        # --- TreatmentRecords table ---
        td = db.CreateTableDef("TreatmentRecords")
        fld = td.CreateField("TreatmentID", dbLong); fld.Attributes = dbAutoIncrField; td.Fields.Append(fld)
        for n, t, s in [("PatientID",dbLong,0),("TreatmentType",dbText,100),("TreatmentDate",dbDate,0),("DoctorName",dbText,80),("Cost",dbCurrency,0),("Notes",dbText,255)]:
            f = td.CreateField(n, t)
            if t == dbText: f.Size = s
            td.Fields.Append(f)
        idx = td.CreateIndex("PrimaryKey"); idx.Primary = True; idx.Fields.Append(idx.CreateField("TreatmentID")); td.Indexes.Append(idx)
        db.TableDefs.Append(td)

        print("[OK] All tables created")

        # --- Relationships ---
        rel = db.CreateRelation("MemberBorrowing", "Members", "Borrowing", 1)
        f = rel.CreateField("MemberID"); f.ForeignName = "MemberID"; rel.Fields.Append(f)
        db.Relations.Append(rel)

        rel = db.CreateRelation("BookBorrowing", "Books", "Borrowing", 1)
        f = rel.CreateField("BookID"); f.ForeignName = "BookID"; rel.Fields.Append(f)
        db.Relations.Append(rel)

        rel = db.CreateRelation("PatientTreatment", "Patients", "TreatmentRecords", 1)
        f = rel.CreateField("PatientID"); f.ForeignName = "PatientID"; rel.Fields.Append(f)
        db.Relations.Append(rel)
        print("[OK] Relationships created")

        # --- Insert Data ---
        books = [
            ("The Great Gatsby","F. Scott Fitzgerald","Fiction","978-0743273565",8.99,1925),
            ("To Kill a Mockingbird","Harper Lee","Fiction","978-0061120084",7.99,1960),
            ("A Brief History of Time","Stephen Hawking","Science","978-0553380163",12.99,1988),
            ("The Art of War","Sun Tzu","Philosophy","978-1599869773",6.99,500),
            ("Becoming","Michelle Obama","Biography","978-1524763138",14.99,2018),
            ("Sapiens","Yuval Noah Harari","History","978-0062316097",11.99,2011),
            ("1984","George Orwell","Fiction","978-0451524935",9.99,1949),
            ("The Lean Startup","Eric Ries","Business","978-0307887894",13.99,2011),
            ("Educated","Tara Westover","Biography","978-0399590504",10.99,2018),
            ("Thinking Fast and Slow","Daniel Kahneman","Psychology","978-0374533557",11.99,2011),
        ]
        for b in books:
            db.Execute(f"INSERT INTO Books (Title,Author,Genre,ISBN,Price,YearPublished) VALUES ('{b[0]}','{b[1]}','{b[2]}','{b[3]}',{b[4]},{b[5]})")

        members = [
            ("Alice","Johnson","alice.j@email.com","07700100001","#2024-01-15#"),
            ("Bob","Smith","bob.s@email.com","07700100002","#2024-02-20#"),
            ("Claire","Davis","claire.d@email.com","07700100003","#2024-03-10#"),
            ("David","Wilson","david.w@email.com","07700100004","#2024-04-05#"),
            ("Eva","Brown","eva.b@email.com","07700100005","#2024-05-18#"),
        ]
        for m in members:
            db.Execute(f"INSERT INTO Members (FirstName,LastName,Email,Phone,MembershipDate) VALUES ('{m[0]}','{m[1]}','{m[2]}','{m[3]}',{m[4]})")

        borrows = [(1,1,"#2025-01-10#","#2025-01-24#"),(1,3,"#2025-02-01#","#2025-02-15#"),
                    (2,2,"#2025-01-15#","#2025-01-29#"),(3,5,"#2025-02-05#","#2025-02-19#"),
                    (4,7,"#2025-01-20#","#2025-02-03#"),(5,4,"#2025-02-10#","#2025-02-24#"),
                    (2,8,"#2025-03-01#","#2025-03-15#"),(3,6,"#2025-03-05#","#2025-03-19#")]
        for bw in borrows:
            db.Execute(f"INSERT INTO Borrowing (MemberID,BookID,BorrowDate,ReturnDate) VALUES ({bw[0]},{bw[1]},{bw[2]},{bw[3]})")

        patients = [
            ("John","Anderson","#1985-03-15#","07700200001","12 Oak Street Edinburgh"),
            ("Mary","Thompson","#1990-07-22#","07700200002","45 Pine Road Glasgow"),
            ("James","Martin","#1978-11-30#","07700200003","8 Elm Avenue Aberdeen"),
            ("Sarah","White","#1995-04-18#","07700200004","23 Birch Lane Dundee"),
            ("Michael","Taylor","#1982-09-05#","07700200005","67 Cedar Close Inverness"),
        ]
        for p in patients:
            db.Execute(f"INSERT INTO Patients (FirstName,LastName,DOB,Phone,Address) VALUES ('{p[0]}','{p[1]}',{p[2]},'{p[3]}','{p[4]}')")

        treatments = [
            (1,"General Checkup","#2025-01-10#","Dr. Campbell",50.00,"Routine annual checkup"),
            (1,"Blood Test","#2025-01-10#","Dr. Campbell",30.00,"Cholesterol panel"),
            (2,"Physiotherapy","#2025-02-05#","Dr. Fraser",75.00,"Lower back pain session 1"),
            (2,"Physiotherapy","#2025-02-19#","Dr. Fraser",75.00,"Lower back pain session 2"),
            (3,"Dental Cleaning","#2025-01-22#","Dr. MacLeod",85.00,"Routine cleaning and polish"),
            (3,"X-Ray","#2025-02-15#","Dr. Stewart",120.00,"Chest X-ray"),
            (4,"Eye Examination","#2025-03-01#","Dr. Murray",60.00,"Vision test and prescription"),
            (5,"General Checkup","#2025-02-28#","Dr. Campbell",50.00,"Pre-employment health check"),
            (5,"Blood Test","#2025-02-28#","Dr. Campbell",30.00,"Full blood count"),
        ]
        for t in treatments:
            db.Execute(f"INSERT INTO TreatmentRecords (PatientID,TreatmentType,TreatmentDate,DoctorName,Cost,Notes) VALUES ({t[0]},'{t[1]}',{t[2]},'{t[3]}',{t[4]},'{t[5]}')")

        print("[OK] All data inserted")

        # --- Queries ---
        db.CreateQueryDef("SortedBooks", "SELECT BookID, Title, Author, Genre, Price FROM Books ORDER BY Genre ASC, Price DESC")
        db.CreateQueryDef("PatientTreatmentSummary",
            "SELECT Patients.PatientID, Patients.FirstName & ' ' & Patients.LastName AS PatientName, "
            "TreatmentRecords.TreatmentType, TreatmentRecords.TreatmentDate, "
            "TreatmentRecords.DoctorName, TreatmentRecords.Cost "
            "FROM Patients INNER JOIN TreatmentRecords ON Patients.PatientID = TreatmentRecords.PatientID "
            "ORDER BY Patients.LastName, TreatmentRecords.TreatmentDate")
        print("[OK] Queries created")

        db.Close()
        access.Quit()
        print("[OK] Phase 1 complete - tables, data, queries")
        return True

    except Exception as e:
        print(f"[ERROR] Phase 1: {e}")
        import traceback; traceback.print_exc()
        try: access.Quit()
        except: pass
        return False
    finally:
        pythoncom.CoUninitialize()


def create_forms():
    """Phase 2: Create forms using Access Application"""
    pythoncom.CoInitialize()
    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False

    try:
        access.OpenCurrentDatabase(ACCDB_PATH)
        time.sleep(1)

        # --- BorrowForm ---
        print("Creating BorrowForm...")
        frm = access.CreateForm()
        frm.RecordSource = "Borrowing"
        frm.Caption = "Library Book Borrowing Form"
        frm.DefaultView = 0
        fn = frm.Name

        # Title
        c = access.CreateControl(fn, 100, 0, "", "LIBRARY BORROWING FORM", 500, 200, 4000, 400)
        c.FontSize = 14; c.FontBold = True

        # BorrowID
        access.CreateControl(fn, 100, 0, "", "Borrow ID:", 500, 800, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtBorrowID", 2200, 800, 2000, 300)
        c.ControlSource = "BorrowID"

        # MemberID
        access.CreateControl(fn, 100, 0, "", "Member ID:", 500, 1300, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtMemberID", 2200, 1300, 2000, 300)
        c.ControlSource = "MemberID"

        # BookID Listbox
        access.CreateControl(fn, 100, 0, "", "Select Book:", 500, 1800, 1500, 300).FontBold = True
        lst = access.CreateControl(fn, 110, 0, "", "lstBooks", 2200, 1800, 3500, 1200)
        lst.RowSourceType = "Table/Query"
        lst.RowSource = "SELECT BookID, Title, Author FROM Books"
        lst.ColumnCount = 3
        lst.ColumnWidths = "500;2000;1500"
        lst.BoundColumn = 1

        # BorrowDate
        access.CreateControl(fn, 100, 0, "", "Borrow Date:", 500, 3200, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtBorrowDate", 2200, 3200, 2000, 300)
        c.ControlSource = "BorrowDate"

        # ReturnDate
        access.CreateControl(fn, 100, 0, "", "Return Date:", 500, 3700, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtReturnDate", 2200, 3700, 2000, 300)
        c.ControlSource = "ReturnDate"

        # Buttons
        b = access.CreateControl(fn, 104, 0, "", "", 500, 4400, 1800, 450); b.Caption = "Save Record"; b.Name = "cmdSave"
        b = access.CreateControl(fn, 104, 0, "", "", 2500, 4400, 1800, 450); b.Caption = "New Record"; b.Name = "cmdNew"
        b = access.CreateControl(fn, 104, 0, "", "", 4500, 4400, 1800, 450); b.Caption = "Close Form"; b.Name = "cmdClose"

        access.DoCmd.Save(2, fn)
        access.DoCmd.Close(2, fn, 1)  # acSaveYes=1
        time.sleep(0.5)
        access.DoCmd.Rename("BorrowForm", 2, fn)
        print("[OK] BorrowForm created")

        # --- TreatmentEntryForm ---
        print("Creating TreatmentEntryForm...")
        frm2 = access.CreateForm()
        frm2.RecordSource = "TreatmentRecords"
        frm2.Caption = "Patient Treatment Entry Form"
        frm2.DefaultView = 0
        fn2 = frm2.Name

        # Title
        c = access.CreateControl(fn2, 100, 0, "", "TREATMENT ENTRY FORM", 500, 200, 4000, 400)
        c.FontSize = 14; c.FontBold = True

        # TreatmentID
        access.CreateControl(fn2, 100, 0, "", "Treatment ID:", 500, 800, 1500, 300).FontBold = True
        c = access.CreateControl(fn2, 109, 0, "", "txtTreatmentID", 2200, 800, 2000, 300)
        c.ControlSource = "TreatmentID"

        # PatientID Listbox
        access.CreateControl(fn2, 100, 0, "", "Select Patient:", 500, 1300, 1500, 300).FontBold = True
        lst2 = access.CreateControl(fn2, 110, 0, "", "lstPatients", 2200, 1300, 3500, 1000)
        lst2.RowSourceType = "Table/Query"
        lst2.RowSource = "SELECT PatientID, FirstName & ' ' & LastName AS PatientName FROM Patients"
        lst2.ColumnCount = 2
        lst2.ColumnWidths = "500;2500"
        lst2.BoundColumn = 1

        # TreatmentType
        access.CreateControl(fn2, 100, 0, "", "Treatment Type:", 500, 2500, 1500, 300).FontBold = True
        c = access.CreateControl(fn2, 109, 0, "", "txtType", 2200, 2500, 2500, 300)
        c.ControlSource = "TreatmentType"

        # TreatmentDate
        access.CreateControl(fn2, 100, 0, "", "Treatment Date:", 500, 3000, 1500, 300).FontBold = True
        c = access.CreateControl(fn2, 109, 0, "", "txtDate", 2200, 3000, 2000, 300)
        c.ControlSource = "TreatmentDate"

        # DoctorName
        access.CreateControl(fn2, 100, 0, "", "Doctor Name:", 500, 3500, 1500, 300).FontBold = True
        c = access.CreateControl(fn2, 109, 0, "", "txtDoctor", 2200, 3500, 2000, 300)
        c.ControlSource = "DoctorName"

        # Cost
        access.CreateControl(fn2, 100, 0, "", "Cost:", 500, 4000, 1500, 300).FontBold = True
        c = access.CreateControl(fn2, 109, 0, "", "txtCost", 2200, 4000, 2000, 300)
        c.ControlSource = "Cost"

        # Notes
        access.CreateControl(fn2, 100, 0, "", "Notes:", 500, 4500, 1500, 300).FontBold = True
        c = access.CreateControl(fn2, 109, 0, "", "txtNotes", 2200, 4500, 3000, 500)
        c.ControlSource = "Notes"

        # Buttons
        b = access.CreateControl(fn2, 104, 0, "", "", 500, 5200, 1800, 450); b.Caption = "Save"; b.Name = "cmdSaveTreatment"
        b = access.CreateControl(fn2, 104, 0, "", "", 2500, 5200, 1800, 450); b.Caption = "New Entry"; b.Name = "cmdNewEntry"
        b = access.CreateControl(fn2, 104, 0, "", "", 4500, 5200, 1800, 450); b.Caption = "Close"; b.Name = "cmdCloseTreatment"

        access.DoCmd.Save(2, fn2)
        access.DoCmd.Close(2, fn2, 1)
        time.sleep(0.5)
        access.DoCmd.Rename("TreatmentEntryForm", 2, fn2)
        print("[OK] TreatmentEntryForm created")

        access.CloseCurrentDatabase()
        print(f"\n[DONE] Access database complete: {ACCDB_PATH}")

    except Exception as e:
        print(f"[ERROR] Phase 2: {e}")
        import traceback; traceback.print_exc()
    finally:
        try: access.Quit()
        except: pass
        time.sleep(1)
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    if create_tables_and_data():
        time.sleep(2)
        gc.collect()
        create_forms()
