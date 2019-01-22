Function clean(s)
  clean = Replace(Replace(Replace(s, "\", "\\"), ",", "\,"), "'", "\'")
End Function
Function win(s)
  win = Replace(s, "/", "\")
End Function
Sub generateData()
Dim sFileLocation, sFileURL, sOutfile, sLink, sJLink, sJTitle, sTitle, initialSource, solrIP As String
Dim nRowsPerPage As Integer
Dim lLastRow As Long
sFileURL = "."
'sOutfile = "gendata.js"
sOutfile = Application.ActiveWorkbook.Path & "/" & "gendata.js"
#If Mac Then
  'sOutfile = Range("Data!B8") & "/" & "gendata.js"
#Else
  'sOutfile = Range("Data!B7") & "/" & "gendata.js"
  sOutfile = win(sOutfile)
#End If
sTitle = Range("Data!B1")
nRowsPerPage = Range("Data!B9")  '29 ' 30 and 860 viewer height for FHD
lLastRow = ActiveSheet.UsedRange.Rows.Count
solrIP = Range("Data!B2")
docRoot = Range("Data!B3")
docRootLength = Len(docRoot)
relDocRoot = Range("Data!B4")
initialSource = Replace(Range("Documents!F2"), "\", "\\")

Open sOutfile For Output As #1

'===========================================================================================================================
'===========================================================================================================================
'BEGIN JAVASCRIPT
Print #1, "maxRows=" & lLastRow - 1 & ";"
Print #1, "var jRowsPerPage=" & nRowsPerPage & "+1;"
Print #1, "jTitle=""" & sTitle & """;"
Print #1, "jInitialSrc=""" & initialSource & """;"


'Declare an in-memory Javascript array ("items[]") to hold values for all table rows - to (re)populate the displayed 29-row table
Print #1, "var items = [];"
Print #1, "var row;"
For i = 2 To lLastRow - 1
  Print #1, "row = {Id:'" & i & "',Date:'" & Range("A" & i) & "',Score:'" & Range("B" & i) & "',Title:'" & clean(Range("C" & i)) & "',DocType:'" & Range("D" & i) & _
    "',Summary:'" & clean(Range("E" & i)) & "',Path:'" & clean(Range("F" & i)) & "'};"
  Print #1, "items.push(row);"
Next i

'Print #1, "alert(items[2].Path)")
Close #1

MsgBox ("Binder Index data file named " & sOutfile & " has been generated.")

End Sub
