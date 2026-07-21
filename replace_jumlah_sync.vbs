Set objWord = CreateObject("Word.Application")
objWord.Visible = False
objWord.DisplayAlerts = 0 ' wdAlertsNone

Sub ReplacePlaceholder(docPath, findText, replaceText)
    Set objDoc = objWord.Documents.Open(docPath)
    
    Set objSelection = objWord.Selection
    objSelection.HomeKey 6 ' wdStory
    objSelection.Find.ClearFormatting
    objSelection.Find.Replacement.ClearFormatting
    
    objSelection.Find.Text = findText
    objSelection.Find.Replacement.Text = replaceText
    objSelection.Find.Forward = True
    objSelection.Find.Wrap = 1 ' wdFindContinue
    objSelection.Find.Format = False
    objSelection.Find.MatchCase = False
    objSelection.Find.MatchWholeWord = False
    
    ' wdReplaceAll = 2
    objSelection.Find.Execute ,,,,,,,,,,2
    
    objDoc.Save
    objDoc.Close False
End Sub

ReplacePlaceholder "g:\Skripsi Sri Darni\Draft_Bab_3_Skripsi.docx", "[Jumlah]", "35"
ReplacePlaceholder "g:\Skripsi Sri Darni\File Skripsi\Sri Darni.docx", "[Jumlah]", "35"

objWord.Quit
WScript.Echo "Done!"
