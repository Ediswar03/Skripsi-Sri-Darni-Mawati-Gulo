Set objWord = CreateObject("Word.Application")
objWord.Visible = False
objWord.DisplayAlerts = 0

Dim docPath
docPath = "g:\Skripsi Sri Darni\File Skripsi\Sri Darni.docx"

Set objDoc = objWord.Documents.Open(docPath)
Set objSel = objWord.Selection

WScript.Echo "Document opened."

' ============================================================
' HELPER: Find text and return Range
' ============================================================
Function FindRange(searchText)
    Dim rng
    Set rng = objDoc.Content
    rng.Find.ClearFormatting
    rng.Find.Text = searchText
    rng.Find.Forward = True
    rng.Find.Wrap = 0
    rng.Find.MatchCase = False
    If rng.Find.Execute() Then
        Set FindRange = rng
    Else
        Set FindRange = Nothing
    End If
End Function

' ============================================================
' REVISION 0: [Jumlah] -> 35
' ============================================================
WScript.Echo "Rev 0: Replace [Jumlah] with 35..."
objSel.HomeKey 6
objSel.Find.ClearFormatting
objSel.Find.Replacement.ClearFormatting
objSel.Find.Text = "[Jumlah]"
objSel.Find.Replacement.Text = "35"
objSel.Find.Forward = True
objSel.Find.Wrap = 1
objSel.Find.Format = False
objSel.Find.MatchCase = False
objSel.Find.MatchWholeWord = False
objSel.Find.Execute ,,,,,,,,,,2
WScript.Echo "  Done."

' ============================================================
' REVISION 1: Insert Hasibuan (2019) before Bahits in A.3.a
' ============================================================
WScript.Echo "Rev 1: Insert Hasibuan (2019) in A.3.a..."

Dim newDef
newDef = "Menurut Hasibuan (2019), kompensasi adalah semua pendapatan yang berbentuk uang, barang langsung atau tidak langsung yang diterima karyawan sebagai imbalan atas jasa yang diberikan kepada organisasi. Kompensasi bertujuan untuk memotivasi pegawai agar lebih giat bekerja dan memberikan pelayanan yang terbaik bagi organisasi."

Dim rngKomp
Set rngKomp = FindRange("Menurut Bahits dkk. (2023:109)")
If Not rngKomp Is Nothing Then
    rngKomp.Collapse 1
    rngKomp.InsertBefore newDef & Chr(13)
    WScript.Echo "  Inserted."
Else
    WScript.Echo "  WARNING: Not found."
End If

' ============================================================
' REVISION 2: Add 3 rows AFTER the table using Tab on last cell
' Tab in last cell of table auto-creates a new row.
' ============================================================
WScript.Echo "Rev 2: Add 3 rows to Tabel 2.2..."

Dim tbl2
Set tbl2 = objDoc.Tables(2)
Dim colCount
colCount = tbl2.Columns.Count
Dim rowCount
rowCount = tbl2.Rows.Count

WScript.Echo "  Cols: " & colCount & ", Rows: " & rowCount

' Select end of last cell and use Tab to create new rows
Dim lastCell
Set lastCell = tbl2.Cell(rowCount, colCount)
lastCell.Range.Collapse 0 ' wdCollapseEnd = 0

' Move selection to end of last cell
lastCell.Range.Select
objSel.Collapse 0 ' move to end
' Press Tab 3 times to create 3 new rows
objSel.TypeText Chr(9) ' Tab -> new row 1
objSel.TypeText Chr(9) ' Tab -> jump to col 2 
objSel.TypeText Chr(9)
objSel.TypeText Chr(9)
objSel.TypeText Chr(9)
objSel.TypeText Chr(9)
' This adds 1 row. Repeat for rows 2 and 3:
' After pressing Tab enough times to reach last col of row we just created, Tab again
' Actually each Tab moves to next cell, and Tab at last cell creates new row.
' We need colCount - 1 more Tabs to reach the last col, then Tab for new row.

' Better: Use InsertRowsBelow on last cell
Set lastCell = tbl2.Cell(tbl2.Rows.Count, colCount)
lastCell.Range.Select
objWord.CommandBars.ExecuteMso "TableInsertRowBelow"
WScript.Echo "  Row insert attempt 1."

Set lastCell = tbl2.Cell(tbl2.Rows.Count, colCount)
lastCell.Range.Select
objWord.CommandBars.ExecuteMso "TableInsertRowBelow"
WScript.Echo "  Row insert attempt 2."

Set lastCell = tbl2.Cell(tbl2.Rows.Count, colCount)
lastCell.Range.Select  
objWord.CommandBars.ExecuteMso "TableInsertRowBelow"
WScript.Echo "  Row insert attempt 3."

rowCount = tbl2.Rows.Count
WScript.Echo "  New row count: " & rowCount

' Now fill last 3 rows
Dim naldRow
naldRow = rowCount - 2

' Row for Naldi
tbl2.Cell(naldRow, 1).Range.Text = "Naldi, M. K., dkk. (2025)"
tbl2.Cell(naldRow, 2).Range.Text = "Pengaruh Kepuasan Kerja, Gaya Kepemimpinan, dan Komitmen Organisasi terhadap Motivasi Kerja"
tbl2.Cell(naldRow, 3).Range.Text = "Motivasi Kerja (Y)"
tbl2.Cell(naldRow, 4).Range.Text = "Gaya Kepemimpinan (X1)"
tbl2.Cell(naldRow, 5).Range.Text = "-"
tbl2.Cell(naldRow, 6).Range.Text = "Gaya kepemimpinan berpengaruh positif dan signifikan terhadap motivasi kerja karyawan."
tbl2.Cell(naldRow, 7).Range.Text = "Perbedaan terletak pada variabel terikat yang hanya meneliti motivasi kerja, bukan kinerja. Penelitian ini tidak menggunakan mediasi."

' Row for Putri
tbl2.Cell(naldRow+1, 1).Range.Text = "Putri, D. R., & Rozi, F. (2024)"
tbl2.Cell(naldRow+1, 2).Range.Text = "Pengaruh Kompensasi dan Lingkungan Kerja terhadap Kinerja Karyawan"
tbl2.Cell(naldRow+1, 3).Range.Text = "Kinerja (Y)"
tbl2.Cell(naldRow+1, 4).Range.Text = "Kompensasi (X1)"
tbl2.Cell(naldRow+1, 5).Range.Text = "-"
tbl2.Cell(naldRow+1, 6).Range.Text = "Kompensasi berpengaruh positif dan signifikan terhadap kinerja karyawan."
tbl2.Cell(naldRow+1, 7).Range.Text = "Perbedaan terletak pada variabel bebas tambahan (lingkungan kerja) dan penelitian ini tidak menggunakan variabel mediasi."

' Row for Satria
tbl2.Cell(naldRow+2, 1).Range.Text = "Satria, E., dkk. (2024)"
tbl2.Cell(naldRow+2, 2).Range.Text = "Pengaruh Kompensasi terhadap Kinerja Pegawai pada Kantor Pemerintahan"
tbl2.Cell(naldRow+2, 3).Range.Text = "Kinerja (Y)"
tbl2.Cell(naldRow+2, 4).Range.Text = "Kompensasi (X1)"
tbl2.Cell(naldRow+2, 5).Range.Text = "-"
tbl2.Cell(naldRow+2, 6).Range.Text = "Kompensasi berpengaruh positif dan signifikan terhadap kinerja pegawai pada instansi pemerintah."
tbl2.Cell(naldRow+2, 7).Range.Text = "Perbedaan terletak pada obyek penelitian instansi pemerintah umum, bukan perangkat desa, dan tidak menggunakan variabel mediasi."

WScript.Echo "  3 rows filled."

' ============================================================
' SAVE
' ============================================================
objDoc.Save
objDoc.Close False
objWord.Quit

WScript.Echo "All revisions done and saved!"
