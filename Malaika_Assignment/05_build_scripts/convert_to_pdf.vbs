Dim word, doc, fso
Set word = CreateObject("Word.Application")
word.Visible = False
Set fso = CreateObject("Scripting.FileSystemObject")

Dim baseDir
baseDir = "C:\Users\admin\Documents\SCQF-L6-Course-Materials\Malaika_Assignment\"

Dim files(1)
files(0) = "Malaika_MGMT268_Assessment1_FINAL.docx"
files(1) = "Malaika_MGMT268_Assessment1_WITH_VIDEO.docx"

Dim i
For i = 0 To 1
    Dim srcPath, pdfPath
    srcPath = baseDir & files(i)
    pdfPath = baseDir & Replace(files(i), ".docx", ".pdf")
    If fso.FileExists(srcPath) Then
        Set doc = word.Documents.Open(srcPath)
        doc.SaveAs2 pdfPath, 17
        doc.Close False
        WScript.Echo "Saved: " & pdfPath
    Else
        WScript.Echo "Not found: " & srcPath
    End If
Next

word.Quit
WScript.Echo "All done."
