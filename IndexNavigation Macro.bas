Attribute VB_Name = "Module1"
Sub CreateSheetIndex()
    Dim ws As Worksheet
    Dim idxWs As Worksheet
    Dim idx As Integer
    Dim firstCell As Range

    ' Delete the Index sheet if it exists
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Index" Then ws.Delete
    Next ws
    Application.DisplayAlerts = True

    ' Create a new Index sheet at the beginning
    Set idxWs = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    idxWs.Name = "Index"

    ' Create the table headers
    idxWs.Range("A1:C1").Value = Array("Serial no.", "Sheet Name", "Sheet Hyperlink")

    idx = 1

    ' Loop through each sheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Index" Then
            ' Insert a row if necessary
            ws.Rows(1).Insert
            ' Clear content in the first cell and unmerge if necessary
            Set firstCell = ws.Cells(1, 1)
            If firstCell.MergeCells Then
                firstCell.UnMerge
            End If
            firstCell.Clear
            firstCell.Interior.Color = RGB(240, 240, 240) ' Set pastel color

            ' Add hyperlink to the first cell, navigating to the "Index" sheet
            ws.Hyperlinks.Add Anchor:=firstCell, Address:="", SubAddress:= _
                "'Index'!A1", TextToDisplay:="Home"

            ' Add to the index
            idxWs.Cells(idx + 1, 1).Value = idx
            idxWs.Cells(idx + 1, 2).Value = ws.Name
            idxWs.Hyperlinks.Add Anchor:=idxWs.Cells(idx + 1, 3), Address:="", SubAddress:= _
                "'" & ws.Name & "'!A1", TextToDisplay:="Go to Sheet"
            idx = idx + 1
        End If
    Next ws

    ' Autofit the columns
    idxWs.Columns("A:C").AutoFit

    ' Show a message box
    MsgBox "Index created successfully!", vbInformation
End Sub


