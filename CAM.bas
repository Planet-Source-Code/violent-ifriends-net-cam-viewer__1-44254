Attribute VB_Name = "CAM"
Public CAMNUM As Integer
Global GlobalUser As String

Public Sub LoadTREE(PathandName As String, Tree As TreeView, Caption As String)
Dim intFileNum As Integer
Dim strList As String
Dim i As Integer
  
intFileNum = FreeFile()
  
Open PathandName For Binary As intFileNum

strList = Space(LOF(intFileNum))

Get intFileNum, 1, strList
Close intFileNum
  
Tree.Nodes.Clear

Set nodex = Tree.Nodes.Add(, , "F", Caption, 1)

varList = Split(strList, vbCrLf)

For i = 0 To UBound(varList)
    If varList(i) <> "" Then
        Set nodex = Tree.Nodes.Add("F", tvwChild, varList(i), varList(i), 2)
        nodex.EnsureVisible
    End If
Next
End Sub
Public Sub SavTREE(PathandName As String, ByVal Tree As TreeView)
Dim intFile As Integer
Dim strList As String
Dim x As Integer

For x = 2 To Tree.Nodes.Count
    strList = strList & Tree.Nodes(x).text & vbCrLf
Next

intFile = FreeFile()

If Dir(PathandName) <> "" Then Kill PathandName
  
Open PathandName For Binary As intFile
Put intFile, 1, strList
Close intFile
End Sub
Public Function GrabItBetween(strString, strFind1, strFind2)
On Error Resume Next
Dim C1, C2, C3, C4
C1 = strString
C2 = InStr(1, C1, strFind1)
C3 = InStr(1, C1, strFind2)
GrabItBetween = Mid(C1, (C2 + Len(strFind1)), C3 - (C2 + Len(strFind1)))
End Function
