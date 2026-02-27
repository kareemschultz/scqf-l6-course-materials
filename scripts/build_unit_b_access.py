"""
UNIT B: F1FJ 12 - Spreadsheet & Database
Part 2: Access Database (.accdb) with tables, relationships, forms, queries
"""
import win32com.client
import os
import time
import pythoncom

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FJ12_Spreadsheet_Database")
ACCDB_PATH = os.path.join(SRC, "F1FJ12_Database.accdb")
os.makedirs(SRC, exist_ok=True)

# DAO constants
dbText = 10
dbLong = 4
dbCurrency = 5
dbDate = 8
dbInteger = 3
dbAutoIncrField = 16
dbRelationDeleteCascade = 4096


def create_access_database():
    pythoncom.CoInitialize()

    # Remove existing file
    if os.path.exists(ACCDB_PATH):
        os.remove(ACCDB_PATH)

    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False

    try:
        # Create new database
        db_engine = access.DBEngine
        workspace = db_engine.Workspaces(0)
        db = workspace.CreateDatabase(
            ACCDB_PATH,
            ";LANGID=0x0409;CP=1252;COUNTRY=0",  # English locale
            64  # dbVersion40
        )

        print("Database created. Adding tables...")

        # ============================================================
        # LIBRARY TABLES (Task A: 1.3.1, 1.3.2)
        # ============================================================

        # Books table
        td_books = db.CreateTableDef("Books")
        fld = td_books.CreateField("BookID", dbLong)
        fld.Attributes = dbAutoIncrField
        td_books.Fields.Append(fld)
        for name, ftype, size in [
            ("Title", dbText, 100),
            ("Author", dbText, 80),
            ("Genre", dbText, 50),
            ("ISBN", dbText, 20),
            ("Price", dbCurrency, 0),
            ("YearPublished", dbInteger, 0),
        ]:
            f = td_books.CreateField(name, ftype)
            if ftype == dbText:
                f.Size = size
            td_books.Fields.Append(f)

        idx = td_books.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Fields.Append(idx.CreateField("BookID"))
        td_books.Indexes.Append(idx)
        db.TableDefs.Append(td_books)

        # Members table
        td_members = db.CreateTableDef("Members")
        fld = td_members.CreateField("MemberID", dbLong)
        fld.Attributes = dbAutoIncrField
        td_members.Fields.Append(fld)
        for name, ftype, size in [
            ("FirstName", dbText, 50),
            ("LastName", dbText, 50),
            ("Email", dbText, 100),
            ("Phone", dbText, 20),
            ("MembershipDate", dbDate, 0),
        ]:
            f = td_members.CreateField(name, ftype)
            if ftype == dbText:
                f.Size = size
            td_members.Fields.Append(f)

        idx = td_members.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Fields.Append(idx.CreateField("MemberID"))
        td_members.Indexes.Append(idx)
        db.TableDefs.Append(td_members)

        # Borrowing table
        td_borrow = db.CreateTableDef("Borrowing")
        fld = td_borrow.CreateField("BorrowID", dbLong)
        fld.Attributes = dbAutoIncrField
        td_borrow.Fields.Append(fld)
        for name, ftype in [
            ("MemberID", dbLong),
            ("BookID", dbLong),
            ("BorrowDate", dbDate),
            ("ReturnDate", dbDate),
        ]:
            f = td_borrow.CreateField(name, ftype)
            td_borrow.Fields.Append(f)

        idx = td_borrow.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Fields.Append(idx.CreateField("BorrowID"))
        td_borrow.Indexes.Append(idx)
        db.TableDefs.Append(td_borrow)

        print("[OK] Library tables created (Books, Members, Borrowing)")

        # ============================================================
        # HEALTHCARE TABLES (Task B: 3.2.1, 3.2.2)
        # ============================================================

        # Patients table
        td_patients = db.CreateTableDef("Patients")
        fld = td_patients.CreateField("PatientID", dbLong)
        fld.Attributes = dbAutoIncrField
        td_patients.Fields.Append(fld)
        for name, ftype, size in [
            ("FirstName", dbText, 50),
            ("LastName", dbText, 50),
            ("DOB", dbDate, 0),
            ("Phone", dbText, 20),
            ("Address", dbText, 150),
        ]:
            f = td_patients.CreateField(name, ftype)
            if ftype == dbText:
                f.Size = size
            td_patients.Fields.Append(f)

        idx = td_patients.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Fields.Append(idx.CreateField("PatientID"))
        td_patients.Indexes.Append(idx)
        db.TableDefs.Append(td_patients)

        # TreatmentRecords table
        td_treat = db.CreateTableDef("TreatmentRecords")
        fld = td_treat.CreateField("TreatmentID", dbLong)
        fld.Attributes = dbAutoIncrField
        td_treat.Fields.Append(fld)
        for name, ftype, size in [
            ("PatientID", dbLong, 0),
            ("TreatmentType", dbText, 100),
            ("TreatmentDate", dbDate, 0),
            ("DoctorName", dbText, 80),
            ("Cost", dbCurrency, 0),
            ("Notes", dbText, 255),
        ]:
            f = td_treat.CreateField(name, ftype)
            if ftype == dbText:
                f.Size = size
            td_treat.Fields.Append(f)

        idx = td_treat.CreateIndex("PrimaryKey")
        idx.Primary = True
        idx.Fields.Append(idx.CreateField("TreatmentID"))
        td_treat.Indexes.Append(idx)
        db.TableDefs.Append(td_treat)

        print("[OK] Healthcare tables created (Patients, TreatmentRecords)")

        # ============================================================
        # RELATIONSHIPS
        # ============================================================
        # Members -> Borrowing
        rel1 = db.CreateRelation("MemberBorrowing", "Members", "Borrowing", 1)
        fld_rel = rel1.CreateField("MemberID")
        fld_rel.ForeignName = "MemberID"
        rel1.Fields.Append(fld_rel)
        db.Relations.Append(rel1)

        # Books -> Borrowing
        rel2 = db.CreateRelation("BookBorrowing", "Books", "Borrowing", 1)
        fld_rel = rel2.CreateField("BookID")
        fld_rel.ForeignName = "BookID"
        rel2.Fields.Append(fld_rel)
        db.Relations.Append(rel2)

        # Patients -> TreatmentRecords
        rel3 = db.CreateRelation("PatientTreatment", "Patients", "TreatmentRecords", 1)
        fld_rel = rel3.CreateField("PatientID")
        fld_rel.ForeignName = "PatientID"
        rel3.Fields.Append(fld_rel)
        db.Relations.Append(rel3)

        print("[OK] Relationships created")

        # ============================================================
        # INSERT DATA
        # ============================================================
        # Books data
        books_data = [
            ("The Great Gatsby", "F. Scott Fitzgerald", "Fiction", "978-0743273565", 8.99, 1925),
            ("To Kill a Mockingbird", "Harper Lee", "Fiction", "978-0061120084", 7.99, 1960),
            ("A Brief History of Time", "Stephen Hawking", "Science", "978-0553380163", 12.99, 1988),
            ("The Art of War", "Sun Tzu", "Philosophy", "978-1599869773", 6.99, 500),
            ("Becoming", "Michelle Obama", "Biography", "978-1524763138", 14.99, 2018),
            ("Sapiens", "Yuval Noah Harari", "History", "978-0062316097", 11.99, 2011),
            ("1984", "George Orwell", "Fiction", "978-0451524935", 9.99, 1949),
            ("The Lean Startup", "Eric Ries", "Business", "978-0307887894", 13.99, 2011),
            ("Educated", "Tara Westover", "Biography", "978-0399590504", 10.99, 2018),
            ("Thinking Fast and Slow", "Daniel Kahneman", "Psychology", "978-0374533557", 11.99, 2011),
        ]

        for b in books_data:
            db.Execute(f"INSERT INTO Books (Title, Author, Genre, ISBN, Price, YearPublished) VALUES ('{b[0]}', '{b[1]}', '{b[2]}', '{b[3]}', {b[4]}, {b[5]})")

        # Members data
        members_data = [
            ("Alice", "Johnson", "alice.j@email.com", "07700100001", "#2024-01-15#"),
            ("Bob", "Smith", "bob.s@email.com", "07700100002", "#2024-02-20#"),
            ("Claire", "Davis", "claire.d@email.com", "07700100003", "#2024-03-10#"),
            ("David", "Wilson", "david.w@email.com", "07700100004", "#2024-04-05#"),
            ("Eva", "Brown", "eva.b@email.com", "07700100005", "#2024-05-18#"),
        ]

        for m in members_data:
            db.Execute(f"INSERT INTO Members (FirstName, LastName, Email, Phone, MembershipDate) VALUES ('{m[0]}', '{m[1]}', '{m[2]}', '{m[3]}', {m[4]})")

        # Borrowing data
        borrows = [
            (1, 1, "#2025-01-10#", "#2025-01-24#"),
            (1, 3, "#2025-02-01#", "#2025-02-15#"),
            (2, 2, "#2025-01-15#", "#2025-01-29#"),
            (3, 5, "#2025-02-05#", "#2025-02-19#"),
            (4, 7, "#2025-01-20#", "#2025-02-03#"),
            (5, 4, "#2025-02-10#", "#2025-02-24#"),
            (2, 8, "#2025-03-01#", "#2025-03-15#"),
            (3, 6, "#2025-03-05#", "#2025-03-19#"),
        ]

        for bw in borrows:
            db.Execute(f"INSERT INTO Borrowing (MemberID, BookID, BorrowDate, ReturnDate) VALUES ({bw[0]}, {bw[1]}, {bw[2]}, {bw[3]})")

        # Patients data
        patients_data = [
            ("John", "Anderson", "#1985-03-15#", "07700200001", "12 Oak Street, Edinburgh"),
            ("Mary", "Thompson", "#1990-07-22#", "07700200002", "45 Pine Road, Glasgow"),
            ("James", "Martin", "#1978-11-30#", "07700200003", "8 Elm Avenue, Aberdeen"),
            ("Sarah", "White", "#1995-04-18#", "07700200004", "23 Birch Lane, Dundee"),
            ("Michael", "Taylor", "#1982-09-05#", "07700200005", "67 Cedar Close, Inverness"),
        ]

        for p in patients_data:
            db.Execute(f"INSERT INTO Patients (FirstName, LastName, DOB, Phone, Address) VALUES ('{p[0]}', '{p[1]}', {p[2]}, '{p[3]}', '{p[4]}')")

        # TreatmentRecords data
        treatments = [
            (1, "General Checkup", "#2025-01-10#", "Dr. Campbell", 50.00, "Routine annual checkup"),
            (1, "Blood Test", "#2025-01-10#", "Dr. Campbell", 30.00, "Cholesterol panel"),
            (2, "Physiotherapy", "#2025-02-05#", "Dr. Fraser", 75.00, "Lower back pain session 1"),
            (2, "Physiotherapy", "#2025-02-19#", "Dr. Fraser", 75.00, "Lower back pain session 2"),
            (3, "Dental Cleaning", "#2025-01-22#", "Dr. MacLeod", 85.00, "Routine cleaning and polish"),
            (3, "X-Ray", "#2025-02-15#", "Dr. Stewart", 120.00, "Chest X-ray"),
            (4, "Eye Examination", "#2025-03-01#", "Dr. Murray", 60.00, "Vision test and prescription update"),
            (5, "General Checkup", "#2025-02-28#", "Dr. Campbell", 50.00, "Pre-employment health check"),
            (5, "Blood Test", "#2025-02-28#", "Dr. Campbell", 30.00, "Full blood count"),
        ]

        for t in treatments:
            db.Execute(f"INSERT INTO TreatmentRecords (PatientID, TreatmentType, TreatmentDate, DoctorName, Cost, Notes) VALUES ({t[0]}, '{t[1]}', {t[2]}, '{t[3]}', {t[4]}, '{t[5]}')")

        print("[OK] All data inserted")

        # ============================================================
        # QUERIES
        # ============================================================

        # Task 1.3.2: SortedBooks query (Genre ASC, Price DESC)
        sorted_books_sql = "SELECT BookID, Title, Author, Genre, Price FROM Books ORDER BY Genre ASC, Price DESC"
        qd1 = db.CreateQueryDef("SortedBooks", sorted_books_sql)
        print("[OK] SortedBooks query created")

        # Task 3.2.2: PatientTreatmentSummary query (join Patients + TreatmentRecords)
        patient_treatment_sql = (
            "SELECT Patients.PatientID, Patients.FirstName & ' ' & Patients.LastName AS PatientName, "
            "TreatmentRecords.TreatmentType, TreatmentRecords.TreatmentDate, "
            "TreatmentRecords.DoctorName, TreatmentRecords.Cost "
            "FROM Patients INNER JOIN TreatmentRecords ON Patients.PatientID = TreatmentRecords.PatientID "
            "ORDER BY Patients.LastName, TreatmentRecords.TreatmentDate"
        )
        qd2 = db.CreateQueryDef("PatientTreatmentSummary", patient_treatment_sql)
        print("[OK] PatientTreatmentSummary query created")

        db.Close()
        print("[OK] Database closed via DAO")

        # ============================================================
        # Now open in Access to create FORMS
        # ============================================================
        print("\nOpening in Access for form creation...")
        access.OpenCurrentDatabase(ACCDB_PATH)
        time.sleep(2)

        # Task 1.3.1: BorrowForm
        borrow_form_code = """
        Dim frm As Form
        Set frm = CreateForm()
        frm.RecordSource = "Borrowing"
        frm.Caption = "Library Book Borrowing Form"

        ' Set form properties
        frm.DefaultView = 0  ' Single Form
        frm.NavigationButtons = True

        ' Add labels and textboxes
        Dim ctl As Control

        ' Title Label
        Set ctl = CreateControl(frm.Name, acLabel, acDetail, , "LIBRARY BORROWING FORM", 500, 200, 4000, 400)
        ctl.FontSize = 14
        ctl.FontBold = True
        ctl.ForeColor = RGB(0, 51, 102)

        ' BorrowID
        Set ctl = CreateControl(frm.Name, acLabel, acDetail, , "Borrow ID:", 500, 800, 1500, 300)
        ctl.FontBold = True
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , "BorrowID", 2200, 800, 2000, 300)

        ' MemberID
        Set ctl = CreateControl(frm.Name, acLabel, acDetail, , "Member ID:", 500, 1300, 1500, 300)
        ctl.FontBold = True
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , "MemberID", 2200, 1300, 2000, 300)

        ' BookID - using a listbox
        Set ctl = CreateControl(frm.Name, acLabel, acDetail, , "Select Book:", 500, 1800, 1500, 300)
        ctl.FontBold = True
        Set ctl = CreateControl(frm.Name, acListBox, acDetail, , "BookID", 2200, 1800, 3500, 1500)
        ctl.RowSourceType = "Table/Query"
        ctl.RowSource = "SELECT BookID, Title, Author FROM Books"
        ctl.ColumnCount = 3
        ctl.ColumnWidths = "500;2000;1500"
        ctl.BoundColumn = 1

        ' BorrowDate
        Set ctl = CreateControl(frm.Name, acLabel, acDetail, , "Borrow Date:", 500, 3500, 1500, 300)
        ctl.FontBold = True
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , "BorrowDate", 2200, 3500, 2000, 300)

        ' ReturnDate
        Set ctl = CreateControl(frm.Name, acLabel, acDetail, , "Return Date:", 500, 4000, 1500, 300)
        ctl.FontBold = True
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , "ReturnDate", 2200, 4000, 2000, 300)

        ' Command Buttons
        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 500, 4700, 1800, 450)
        ctl.Caption = "Save Record"
        ctl.Name = "cmdSave"
        ctl.OnClick = "=DoCmd.RunCommand(acCmdSaveRecord)"

        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 2500, 4700, 1800, 450)
        ctl.Caption = "New Record"
        ctl.Name = "cmdNew"

        Set ctl = CreateControl(frm.Name, acCommandButton, acDetail, , , 4500, 4700, 1800, 450)
        ctl.Caption = "Close Form"
        ctl.Name = "cmdClose"

        DoCmd.Save acForm, frm.Name
        DoCmd.Rename "BorrowForm", acForm, frm.Name
        DoCmd.Close acForm, "BorrowForm"
        """

        print("Creating BorrowForm...")
        try:
            frm = access.CreateForm()
            frm.RecordSource = "Borrowing"
            frm.Caption = "Library Book Borrowing Form"
            frm.DefaultView = 0  # Single Form

            form_name = frm.Name

            # Title label
            ctl = access.CreateControl(form_name, 100, 0, "", "LIBRARY BORROWING FORM", 500, 200, 4000, 400)
            ctl.FontSize = 14
            ctl.FontBold = True
            ctl.ForeColor = 0x003366

            # BorrowID label + textbox
            ctl = access.CreateControl(form_name, 100, 0, "", "Borrow ID:", 500, 800, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name, 109, 0, "", "BorrowID", 2200, 800, 2000, 300)
            ctl.ControlSource = "BorrowID"

            # MemberID
            ctl = access.CreateControl(form_name, 100, 0, "", "Member ID:", 500, 1300, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name, 109, 0, "", "MemberID", 2200, 1300, 2000, 300)
            ctl.ControlSource = "MemberID"

            # BookID - Listbox
            ctl = access.CreateControl(form_name, 100, 0, "", "Select Book:", 500, 1800, 1500, 300)
            ctl.FontBold = True
            lst = access.CreateControl(form_name, 110, 0, "", "BookList", 2200, 1800, 3500, 1200)
            lst.RowSourceType = "Table/Query"
            lst.RowSource = "SELECT BookID, Title, Author FROM Books"
            lst.ColumnCount = 3
            lst.ColumnWidths = "500;2000;1500"
            lst.BoundColumn = 1

            # BorrowDate
            ctl = access.CreateControl(form_name, 100, 0, "", "Borrow Date:", 500, 3200, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name, 109, 0, "", "BorrowDate", 2200, 3200, 2000, 300)
            ctl.ControlSource = "BorrowDate"

            # ReturnDate
            ctl = access.CreateControl(form_name, 100, 0, "", "Return Date:", 500, 3700, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name, 109, 0, "", "ReturnDate", 2200, 3700, 2000, 300)
            ctl.ControlSource = "ReturnDate"

            # Command buttons
            btn1 = access.CreateControl(form_name, 104, 0, "", "", 500, 4400, 1800, 450)
            btn1.Caption = "Save Record"
            btn1.Name = "cmdSave"

            btn2 = access.CreateControl(form_name, 104, 0, "", "", 2500, 4400, 1800, 450)
            btn2.Caption = "New Record"
            btn2.Name = "cmdNew"

            btn3 = access.CreateControl(form_name, 104, 0, "", "", 4500, 4400, 1800, 450)
            btn3.Caption = "Close Form"
            btn3.Name = "cmdClose"

            access.DoCmd.Save(2, form_name)  # acForm = 2
            access.DoCmd.Close(2, form_name)
            access.DoCmd.Rename("BorrowForm", 2, form_name)
            print("[OK] BorrowForm created")

        except Exception as e:
            print(f"[WARN] BorrowForm creation: {e}")

        # Create TreatmentEntryForm
        try:
            frm2 = access.CreateForm()
            frm2.RecordSource = "TreatmentRecords"
            frm2.Caption = "Patient Treatment Entry Form"
            frm2.DefaultView = 0

            form_name2 = frm2.Name

            # Title
            ctl = access.CreateControl(form_name2, 100, 0, "", "TREATMENT ENTRY FORM", 500, 200, 4000, 400)
            ctl.FontSize = 14
            ctl.FontBold = True
            ctl.ForeColor = 0x003366

            # TreatmentID
            ctl = access.CreateControl(form_name2, 100, 0, "", "Treatment ID:", 500, 800, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name2, 109, 0, "", "TreatmentID", 2200, 800, 2000, 300)
            ctl.ControlSource = "TreatmentID"

            # PatientID - Listbox with patient names
            ctl = access.CreateControl(form_name2, 100, 0, "", "Select Patient:", 500, 1300, 1500, 300)
            ctl.FontBold = True
            lst2 = access.CreateControl(form_name2, 110, 0, "", "PatientList", 2200, 1300, 3500, 1000)
            lst2.RowSourceType = "Table/Query"
            lst2.RowSource = "SELECT PatientID, FirstName & ' ' & LastName AS PatientName FROM Patients"
            lst2.ColumnCount = 2
            lst2.ColumnWidths = "500;2500"
            lst2.BoundColumn = 1

            # TreatmentType
            ctl = access.CreateControl(form_name2, 100, 0, "", "Treatment Type:", 500, 2500, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name2, 109, 0, "", "TreatmentType", 2200, 2500, 2500, 300)
            ctl.ControlSource = "TreatmentType"

            # TreatmentDate
            ctl = access.CreateControl(form_name2, 100, 0, "", "Treatment Date:", 500, 3000, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name2, 109, 0, "", "TreatmentDate", 2200, 3000, 2000, 300)
            ctl.ControlSource = "TreatmentDate"

            # DoctorName
            ctl = access.CreateControl(form_name2, 100, 0, "", "Doctor Name:", 500, 3500, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name2, 109, 0, "", "DoctorName", 2200, 3500, 2000, 300)
            ctl.ControlSource = "DoctorName"

            # Cost
            ctl = access.CreateControl(form_name2, 100, 0, "", "Cost (£):", 500, 4000, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name2, 109, 0, "", "Cost", 2200, 4000, 2000, 300)
            ctl.ControlSource = "Cost"

            # Notes
            ctl = access.CreateControl(form_name2, 100, 0, "", "Notes:", 500, 4500, 1500, 300)
            ctl.FontBold = True
            ctl = access.CreateControl(form_name2, 109, 0, "", "Notes", 2200, 4500, 3000, 600)
            ctl.ControlSource = "Notes"

            # Command buttons
            btn = access.CreateControl(form_name2, 104, 0, "", "", 500, 5300, 1800, 450)
            btn.Caption = "Save Treatment"
            btn.Name = "cmdSaveTreatment"

            btn = access.CreateControl(form_name2, 104, 0, "", "", 2500, 5300, 1800, 450)
            btn.Caption = "New Entry"
            btn.Name = "cmdNewEntry"

            btn = access.CreateControl(form_name2, 104, 0, "", "", 4500, 5300, 1800, 450)
            btn.Caption = "Close"
            btn.Name = "cmdCloseTreatment"

            access.DoCmd.Save(2, form_name2)
            access.DoCmd.Close(2, form_name2)
            access.DoCmd.Rename("TreatmentEntryForm", 2, form_name2)
            print("[OK] TreatmentEntryForm created")

        except Exception as e:
            print(f"[WARN] TreatmentEntryForm creation: {e}")

        access.CloseCurrentDatabase()
        print(f"\n[DONE] Access database saved: {ACCDB_PATH}")

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            access.Quit()
        except:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    create_access_database()
