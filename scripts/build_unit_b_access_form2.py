"""
Create TreatmentEntryForm in existing Access database
(Separate session to avoid COM server crash from creating 2 forms in one session)
"""
import win32com.client
import os
import time
import pythoncom
import subprocess

BASE = r"C:\Users\admin\Documents\SCQF-L6-Course-Materials\SCQF_L6_FINAL_SUBMISSION"
SRC = os.path.join(BASE, "source_files", "F1FJ12_Spreadsheet_Database")
ACCDB_PATH = os.path.join(SRC, "F1FJ12_Database.accdb")


def create_treatment_form():
    pythoncom.CoInitialize()

    # Kill any existing Access
    subprocess.run(['taskkill', '/F', '/IM', 'MSACCESS.EXE'], capture_output=True)
    time.sleep(2)

    access = win32com.client.Dispatch("Access.Application")
    access.Visible = False

    try:
        access.OpenCurrentDatabase(ACCDB_PATH)
        time.sleep(1)

        print("Creating TreatmentEntryForm...")
        frm = access.CreateForm()
        frm.RecordSource = "TreatmentRecords"
        frm.Caption = "Patient Treatment Entry Form"
        frm.DefaultView = 0
        fn = frm.Name

        # Title
        c = access.CreateControl(fn, 100, 0, "", "TREATMENT ENTRY FORM", 500, 200, 4000, 400)
        c.FontSize = 14
        c.FontBold = True

        # TreatmentID
        access.CreateControl(fn, 100, 0, "", "Treatment ID:", 500, 800, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtTreatmentID", 2200, 800, 2000, 300)
        c.ControlSource = "TreatmentID"

        # PatientID Listbox
        access.CreateControl(fn, 100, 0, "", "Select Patient:", 500, 1300, 1500, 300).FontBold = True
        lst = access.CreateControl(fn, 110, 0, "", "lstPatients", 2200, 1300, 3500, 1000)
        lst.RowSourceType = "Table/Query"
        lst.RowSource = "SELECT PatientID, FirstName & ' ' & LastName AS PatientName FROM Patients"
        lst.ColumnCount = 2
        lst.ColumnWidths = "500;2500"
        lst.BoundColumn = 1

        # TreatmentType
        access.CreateControl(fn, 100, 0, "", "Treatment Type:", 500, 2500, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtType", 2200, 2500, 2500, 300)
        c.ControlSource = "TreatmentType"

        # TreatmentDate
        access.CreateControl(fn, 100, 0, "", "Treatment Date:", 500, 3000, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtDate", 2200, 3000, 2000, 300)
        c.ControlSource = "TreatmentDate"

        # DoctorName
        access.CreateControl(fn, 100, 0, "", "Doctor Name:", 500, 3500, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtDoctor", 2200, 3500, 2000, 300)
        c.ControlSource = "DoctorName"

        # Cost
        access.CreateControl(fn, 100, 0, "", "Cost:", 500, 4000, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtCost", 2200, 4000, 2000, 300)
        c.ControlSource = "Cost"

        # Notes
        access.CreateControl(fn, 100, 0, "", "Notes:", 500, 4500, 1500, 300).FontBold = True
        c = access.CreateControl(fn, 109, 0, "", "txtNotes", 2200, 4500, 3000, 500)
        c.ControlSource = "Notes"

        # Buttons
        b = access.CreateControl(fn, 104, 0, "", "", 500, 5200, 1800, 450)
        b.Caption = "Save"
        b.Name = "cmdSaveTreatment"

        b = access.CreateControl(fn, 104, 0, "", "", 2500, 5200, 1800, 450)
        b.Caption = "New Entry"
        b.Name = "cmdNewEntry"

        b = access.CreateControl(fn, 104, 0, "", "", 4500, 5200, 1800, 450)
        b.Caption = "Close"
        b.Name = "cmdCloseTreatment"

        access.DoCmd.Save(2, fn)
        access.DoCmd.Close(2, fn, 1)
        time.sleep(0.5)
        access.DoCmd.Rename("TreatmentEntryForm", 2, fn)
        print("[OK] TreatmentEntryForm created")

        access.CloseCurrentDatabase()
        print("[DONE] Access database complete")

    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        try:
            access.Quit()
        except:
            pass
        time.sleep(1)
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    create_treatment_form()
