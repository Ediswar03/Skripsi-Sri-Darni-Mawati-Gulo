Set objWord = CreateObject("Word.Application")
objWord.Visible = False
objWord.DisplayAlerts = 0

docPath = "g:\Skripsi Sri Darni\File Skripsi\Sri Darni.docx"
Set objDoc = objWord.Documents.Open(docPath)
Set objSel = objWord.Selection

WScript.Echo "Opened document."

' Find "BAB I" paragraph to place Section Break before it
Set rngBab1 = objDoc.Content
rngBab1.Find.ClearFormatting
rngBab1.Find.Text = "BAB I"
rngBab1.Find.Forward = True
rngBab1.Find.Wrap = 0 ' wdFindStop
rngBab1.Find.MatchCase = True

If rngBab1.Find.Execute() Then
    WScript.Echo "Found BAB I at position " & rngBab1.Start
    
    ' Move selection to beginning of BAB I
    rngBab1.Collapse 1 ' wdCollapseStart
    rngBab1.Select
    
    ' Insert Section Break Next Page (wdSectionBreakNextPage = 2)
    objSel.InsertBreak 2
    WScript.Echo "Section break inserted before BAB I."
Else
    WScript.Echo "WARNING: Could not find BAB I heading."
End If

' Configure Section 1 (Front Matter: Cover, Abstrak, Kata Pengantar, Daftar Isi/Tabel/Gambar)
WScript.Echo "Configuring Section 1 (Roman Numerals)..."
Set sec1 = objDoc.Sections(1)
sec1.PageSetup.DifferentFirstPageHeaderFooter = True ' Cover has no header/footer

' Add page number to Section 1 footer (center, lowercase roman)
' wdHeaderFooterPrimary = 1
Set ftr1 = sec1.Footers(1)
ftr1.PageNumbers.NumberStyle = 2 ' wdPageNumberStyleLowercaseRoman (i, ii, iii...)
ftr1.PageNumbers.HeadingLevelForChapter = 0
ftr1.PageNumbers.IncludeChapterNumber = False
ftr1.PageNumbers.RestartNumberingAtSection = True
ftr1.PageNumbers.StartingNumber = 1
' Add page number at center (wdAlignPageNumberCenter = 1)
ftr1.PageNumbers.Add 1, True

' Configure Section 2 (Body: Bab I to Bab III & Daftar Pustaka)
If objDoc.Sections.Count >= 2 Then
    WScript.Echo "Configuring Section 2 (Arabic Numerals)..."
    Set sec2 = objDoc.Sections(2)
    
    ' Unlink footer from Section 1
    Set ftr2 = sec2.Footers(1)
    ftr2.LinkToPrevious = False
    
    ftr2.PageNumbers.NumberStyle = 0 ' wdPageNumberStyleArabic (1, 2, 3...)
    ftr2.PageNumbers.HeadingLevelForChapter = 0
    ftr2.PageNumbers.IncludeChapterNumber = False
    ftr2.PageNumbers.RestartNumberingAtSection = True
    ftr2.PageNumbers.StartingNumber = 1
    ' Add page number at center
    ftr2.PageNumbers.Add 1, True
    WScript.Echo "Section 2 page numbers configured."
End If

objDoc.Save
objDoc.Close False
objWord.Quit

WScript.Echo "Page numbering and Section Breaks configured successfully!"
