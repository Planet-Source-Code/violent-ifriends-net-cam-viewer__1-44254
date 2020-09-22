Attribute VB_Name = "modStuff"
Public Function SaveLW(Lw As ListView, Fname As String)
    
    Dim FileId As Integer
    Dim X As Integer
    Dim sIdx As Integer
    sIdx = Lw.ColumnHeaders.Count - 1
    FileId = FreeFile
    On Error Resume Next
    Open Fname For Output As #FileId
    For i = 1 To Lw.ListItems.Count
        Write #FileId, Lw.ListItems.Item(i).Text
        For X = 1 To sIdx
            Write #FileId, Lw.ListItems.Item(i).SubItems(X)
        Next
    Next
    Close #FileId
End Function
Public Function LoadLW(Lw As ListView, Fname As String)
    
    Dim FileId As Integer
    Dim LVI As ListItem
    Dim fData As New Collection
    Dim Buffer As String
    Dim X As Integer
    Dim i As Integer
    Dim sIdx As Integer
    sIdx = Lw.ColumnHeaders.Count - 1
    i = 0
    FileId = FreeFile
    On Error Resume Next
    Open Fname For Input As #FileId
    While Not EOF(FileId)
        i = i + 1
        For X = 0 To sIdx
            Input #FileId, Buffer
            fData.Add Buffer
        Next
        Set LVI = Lw.ListItems.Add
        LVI.Text = fData.Item(fData.Count - sIdx)
        For y = 1 To sIdx
            LVI.SubItems(y) = fData.Item(fData.Count + y - sIdx)
        Next
    Wend
    Close #FileId

End Function
Public Function PrintLW(Lw As ListView)
    
    Dim sData As String
    Dim X As Integer
    Dim sIdx As Integer
    sIdx = Lw.ColumnHeaders.Count - 1
    On Error Resume Next
    Printer.Print Tab(4), "ListView print demo"
    Printer.Print
    For i = 1 To Lw.ListItems.Count
        sData = Lw.ListItems.Item(i).Text
        For X = 1 To sIdx
            sData = sData & " " & Lw.ListItems.Item(i).SubItems(X)
        Next
        Printer.Print sData
        sData = ""
    Next
    Printer.EndDoc
    Lw.ListItems.Clear
    MsgBox "Done!!"

End Function

Public Function SaveCombo(cbo As ComboBox, Optional sFilename As String = "\test.txt")
    Dim i As Integer
    Dim X As Integer
    X = FreeFile
    'By default, the file is located in this applications folder with name Test.txt
    '(That is if you don't specify your file name  while calling this function)
    'If you specify the filename, we will use that instead of default name.
    Open App.Path & sFilename For Output As #X
        For i = 0 To cbo.ListCount
            Print #X, cbo.List(i)
        Next i
    Close #X
    cbo.Clear
    MsgBox "Done!"
End Function
Public Function PrintCombo(cbo As ComboBox)
    Dim i As Integer
    Printer.Print Tab(4), "ComboBox print demo"
    Printer.Print
    For i = 0 To cbo.ListCount
        Printer.Print cbo.List(i)
    Next i
    Printer.EndDoc
    cbo.Clear
    MsgBox "Done!"
End Function

Public Function LoadCombo(cbo As ComboBox, Optional sFilename As String = "\test.txt")
    Dim i As Integer
    Dim X As Integer
    X = FreeFile
    Open App.Path & sFilename For Input As #X
    Do While Not EOF(X)
        Input #X, sData
        cbo.AddItem sData
    Loop
    Close #X
    MsgBox "Items loaded!"
End Function

Public Function SaveL(aList As ListBox, sPath As String)
    Dim Nbr As Long
    On Error Resume Next
    Open sPath For Output As #1
    For Nbr = 0 To aList.ListCount - 1
        Print #1, aList.List(Nbr)
    Next Nbr
    Close #1

End Function
Public Function PrintL(aList As ListBox)
    Dim Nbr As Long
    Printer.Print Tab(4), "ListBox print demo"
    Printer.Print
    For Nbr = 0 To aList.ListCount - 1
        Printer.Print aList.List(Nbr)
    Next Nbr
    Printer.EndDoc

End Function

Public Function LoadL(aList As ListBox, sPath As String)

    Dim sText As String
    Dim X As Integer
    X = FreeFile
    On Error Resume Next
    Open sPath For Input As #X
    While Not EOF(X)
        Input #X, sText$
            aList.AddItem sText$
            DoEvents
    Wend
    Close #X

End Function

Public Function SaveTxtBoxes(frmForm As Form, sFile As String)
    Dim sTxt As Object
    Dim sTag As String
    Dim X As Integer
    X = FreeFile
    Open App.Path & sFile For Output As #X
    'I have hardcoded that T as Tag value for each TextBoxes, but
    'you could change it or add it as a parameter for this function...
    For Each sTxt In frmForm.Controls
        If sTxt.Tag = "T" Then
            Write #X, sTxt.Text, sTxt.Name
        End If
        DoEvents
    Next sTxt
    Close #X
    'You can remove this
    MsgBox "Done!"
End Function

Public Function EmbtyText(frm As Form)
    Dim i As Integer
    For i = 0 To frm.Controls.Count - 1
        If TypeOf frm.Controls(i) Is TextBox Then frm.Controls(i) = ""
    Next i
End Function

Public Function LoadTxtBoxes(sFrm As Form, sFile As String)
    Dim X As Integer
    Dim CC As Object
    Dim sText As String
    Dim sTag As String
    X = FreeFile
    Open App.Path & sFile For Input As #X
    Do While Not EOF(X)
        Input #X, sText, sTag
        For i = 0 To sFrm.Controls.Count - 1
            If TypeOf sFrm.Controls(i) Is TextBox Then
                If sFrm.Controls(i).Name = sTag Then
                    sFrm.Controls(i).Text = sText
                End If
            End If
        Next i
    Loop
    Close #X
    'You can remove this
    MsgBox ("OK")
End Function

Public Function PrintTxtBoxes(frmForm As Form)
    Dim sTxt As Object
    Printer.Print Tab(4), "TextBox print demo"
    Printer.Print
    For Each sTxt In frmForm.Controls
        If sTxt.Tag = "T" Then
            Printer.Print sTxt.Name, ": ", sTxt.Text
        End If
        DoEvents
    Next sTxt
    Printer.EndDoc
    'You can remove this
    MsgBox "Done!"
End Function

