Attribute VB_Name = "wizdum"
Global thURL As String

Public Type qHttp
    tpage As String
    trefer As String
    thost As String
    tContent As String
    ReqType As String
    tServer As String
    tPort As String
    tcookie As String
    FilterWatchStart As String
    FilterWatchStartAdd As String
    FilterWatchEnd As String
    FilterWatchEndAdd As String
End Type



Public Sub GetHeader(tpage As String, thost, sckWSC As Winsock)
With sckWSC
    .SendData "GET /" & tpage$ & " HTTP/1.1" & vbCrLf
    .SendData "Accept: text/plain" & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & vbCrLf
    .SendData "Host: " & thost & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf & vbCrLf
End With
End Sub
Public Sub GetHeaderF(tpage As String, trefer As String, thost, sckWSC As Winsock)
'i found this sub on psc back in the day
With sckWSC
    .SendData "GET /" & tpage$ & " HTTP/1.1" & vbCrLf
    .SendData "Accept: application/vnd.ms-excel, application/msword, application/vnd.ms-powerpoint, application/x-comet, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf
    .SendData "Referer: " & trefer & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & vbCrLf
    .SendData "Host: " & thost & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf & vbCrLf
End With
End Sub


Public Sub GetHeaderCH(tpage As String, trefer As String, thost, sckWSC As Winsock)
On Error Resume Next
With sckWSC
    .SendData "GET /" & tpage$ & " HTTP/1.1" & vbCrLf
    .SendData "Accept: */*" & vbCrLf
    .SendData "Accept: application/vnd.ms-excel, application/msword, application/vnd.ms-powerpoint, application/x-comet, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf
    .SendData "Referer: " & trefer & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & vbCrLf
    .SendData "Host: " & thost & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf
End With
End Sub

Public Sub GetHeaderC(tpage As String, tcookie As String, thost, sckWSC As Winsock)
'i found this sub on psc back in the day
With sckWSC
    .SendData "GET /" & tpage$ & " HTTP/1.1" & vbCrLf
    .SendData "Accept: text/plain" & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" & vbCrLf
    .SendData "Host: " & thost & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf
    .SendData "Cookie: " & tcookie & vbCrLf & vbCrLf
End With
End Sub



Public Sub PostHeader(tpage As String, trefer As String, thost, tContent As String, sckWSC As Winsock)
'i made this when i was thinking about creating an account master
On Error Resume Next
With sckWSC
    .SendData "POST " & tpage & " HTTP/1.1" & vbCrLf
    .SendData "Accept: application/vnd.ms-excel, application/msword, application/vnd.ms-powerpoint, application/x-comet, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf
    .SendData "Referer: " & trefer & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; AtHome0102; DigExt)" & vbCrLf
    .SendData "Host: " & thost & vbCrLf
    .SendData "Content-Length: " & Len(tContent) & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf
    .SendData "Cache-Control: no-cache" & vbCrLf & vbCrLf
    .SendData tContent & vbCrLf
End With
End Sub

Public Sub PostPocket(tpage As String, tContent As String, sckWSC As Winsock)
'pocket
With sckWSC
    .SendData "POST " & tpage & " HTTP/1.0" & vbCrLf
    .SendData "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    .SendData "Content-Length: " & Len(tContent) & vbCrLf & vbCrLf
    .SendData tContent & vbCrLf
End With
End Sub

Public Sub PostHeaderC(tpage As String, trefer As String, thost, tcookie As String, tContent As String, sckWSC As Winsock)
'if they require a cookie u can use this biotch
With sckWSC
    .SendData "POST " & tpage & " HTTP/1.1" & vbCrLf
    .SendData "Accept: application/vnd.ms-excel, application/msword, application/vnd.ms-powerpoint, application/x-comet, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf
    .SendData "Referer: " & trefer & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" & vbCrLf
    .SendData "Host: " & thost & vbCrLf
    .SendData "Content-Length: " & Len(tContent) & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf
    .SendData "Cache-Control: no-cache" & vbCrLf
    .SendData "Cookie: " & tcookie & vbCrLf & vbCrLf
    .SendData tContent & vbCrLf
End With
End Sub
Public Sub PostHeaderCC(tpage As String, trefer As String, thost, tcookie As String, tContent As String, sckWSC As Winsock)
'if they require a cookie u can use this biotch
With sckWSC
    .SendData "POST " & tpage & " HTTP/1.1" & vbCrLf
    .SendData "Accept: application/vnd ms-excel, application/msword, application/vnd ms-powerpoint, application/x-comet, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf
    .SendData "Referer: " & trefer & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Content-Type: application/x-www-form-urlencoded" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate"
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" & vbCrLf
    .SendData "Host: " & thost & vbCrLf
    .SendData "Content-Length: " & Len(tContent) & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf
    .SendData "Cache-Control: no-cache" & vbCrLf
    .SendData "Cookie: " & tcookie & vbCrLf & vbCrLf
    .SendData tContent & vbCrLf
End With
End Sub



Public Sub TreeLoot(tX, tY, tProfile, sckWSC As Winsock)
'treeloot.com
With sckWSC
    .SendData "GET /cgi-bin/moneytree.cgi?" & tX & "," & tY & " HTTP/1.1" & vbCrLf
    .SendData "Accept: application/vnd.ms-excel, application/msword, application/vnd.ms-powerpoint, application/x-comet, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf
    .SendData "Referer: http://www.treeloot.com/play/moneytree_spring_grid.html" & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & vbCrLf
    .SendData "Host: game.treeloot.com" & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf
    .SendData "Cookie: profileID=" & tProfile & vbCrLf & vbCrLf
End With
End Sub


Function GetInfo(Start, str, Find)
Dim sName, Over, sChr
sName = ""
Over = 1
Do Until sChr = Find
    sChr = Mid(str, Start + Over, Len(Find))
    Addit = Mid(str, Start + Over, 1)
    Over = Over + 1
    sName = sName + Addit
    If Over > Len(str) Then Exit Function 'time out
  DoEvents
Loop
GetInfo = Left(sName, Len(sName) - 1)
End Function


Function tCounter(tpage, Counter, sckWSC As Winsock)
With sckWSC
    .SendData "GET " & Counter & " HTTP/1.1" & vbCrLf
    .SendData "Accept: */*" & vbCrLf
    .SendData "Referer: " & tpage & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "If-Modified-Since: Thu, 14 Dec 2000 19:12:18 GMT; length=62" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & vbCrLf
    .SendData "Host: angelfire.lycos.com" & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf & vbCrLf
End With
End Function



Public Function GetR(LowerBound, UpperBound)
Randomize
GetR = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function
Function tossIt(Odds)
tossIt = GetR(0, Odds)
End Function

Function GenyChar(Letters, Numbers, Length)
Dim zChar
zChar = vbNullString
If Letters = 1 And Numbers = 1 Then
    For i = 1 To Length
        a = tossIt(1)
        If a = 0 Then
            zChar = zChar & GetR(0, 9)
            Else
            zChar = zChar & Chr(GetR(65, 90))
        End If
    Next i
    ElseIf Letters = 1 Then
        For i = 1 To Length
            zChar = zChar & GetR(0, 9)
        Next
        
    ElseIf Numbers = 1 Then
        For i = 1 To Length
            zChar = zChar & Chr(GetR(65, 90))
        Next
End If
GenyChar = zChar
End Function
Public Function GetRandomChar(dMaxChars As Long, Optional bRandCharCount As Boolean) As String
    
    On Error GoTo EH
    
    Dim MyByteArray() As Byte
    Dim dMyElements As Double
    Dim dCount As Double
    
    Call Randomize
   
    If bRandCharCount Then
 
        dMyElements = (1 + Int(Rnd * dMaxChars)) * 2
    Else

        dMyElements = dMaxChars * 2
    End If
    
    If dMyElements > 0 And dMyElements <= 2 ^ 31 Then

        ReDim MyByteArray(1 To dMyElements) As Byte
    Else
        If dMyElements > 0 Then
            GoTo EH
        Else
            GetRandomChar = vbNullString
            Exit Function
        End If
    End If
    

    For dCount = 1 To dMyElements
        MyByteArray(dCount) = 1 + Int(Rnd * 255)
        dCount = dCount + 1
        MyByteArray(dCount) = 0
    Next
  
    GetRandomChar = MyByteArray
    
    Exit Function
EH:
    MsgBox "The number you have entered is not a whole number. (No Decimals !) " & vbCrLf & _
           "Or you have entered a very Large number.  (" & dMaxChars & ")" & vbCrLf & vbCrLf & _
           "Error Number " & Err.Number & vbCrLf & vbCrLf & Err.Description
    GetRandomChar = vbNullString
End Function

Function Post(thePage, theReferer, theHost, theCookie, theContent)
Post = ""
Post = Post & "POST " & thePage & " HTTP/1.0" & vbCrLf
Post = Post & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/msword, application/vnd.ms-powerpoint, application/vnd.ms-excel, */*" & vbCrLf
If theReferer <> "" Then Post = Post & "Referer: " & theReferer & vbCrLf
Post = Post & "Accept -Language: en -us" & vbCrLf
Post = Post & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
Post = Post & "Accept -Encoding: gzip , deflate" & vbCrLf
Post = Post & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.01; Windows NT; MSOCD; AtHome020)" & vbCrLf
If theHost <> "" Then Post = Post & "Host: " & theHost & vbCrLf
Post = Post & "Content-Length: " & Len(theContent) & vbCrLf
Post = Post & "Proxy -Connection: Keep -Alive" & vbCrLf
Post = Post & "Pragma: no -cache" & vbCrLf
If theCookie <> "" Then Post = Post & "Cookie: " & theCookie & vbCrLf
Post = Post & "" & vbCrLf
Post = Post & theContent & vbCrLf
End Function

Function Gett(thePage, theReferer, theHost, theCookie)
Gett = ""
Gett = Gett & "GET " & thePage & " HTTP/1.0" & vbCrLf
Gett = Gett & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/msword, application/vnd.ms-powerpoint, application/vnd.ms-excel, */*" & vbCrLf
If theReferer <> "" Then Gett = Gett & "Referer: " & theReferer & vbCrLf
Gett = Gett & "Accept -Language: en -us" & vbCrLf
Gett = Gett & "Accept -Encoding: gzip , deflate" & vbCrLf
Gett = Gett & "User-Agent: Mozilla/4.0 (compatible; MSIE 5.01; Windows NT; MSOCD; AtHome020)" & vbCrLf
If theHost <> "" Then Gett = Gett & "Host: " & theHost & vbCrLf
Gett = Gett & "Proxy -Connection: Keep -Alive" & vbCrLf
Gett = Gett & "Pragma: no -cache" & vbCrLf
If theCookie <> "" Then Gett = Gett & "Cookie: " & theCookie & vbCrLf
Gett = Gett & "" & vbCrLf
End Function

Function AIMGet(thePage, theReferer, theHost)
AIMGet = ""
AIMGet = AIMGet & "GET " & thePage & " HTTP/1.0" & vbCrLf
If theHost <> "" Then AIMGet = AIMGet & "Host: " & theHost & vbCrLf
AIMGet = AIMGet & "If-Modified-Since:" & vbCrLf
AIMGet = AIMGet & "Accept:*/*" & vbCrLf
If theReferer <> "" Then AIMGet = AIMGet & "Referer: " & theReferer & vbCrLf
AIMGet = AIMGet & "User-Agent: AIM/30 (Mozilla 1.24b; Windows; I; 32-bit)" & vbCrLf
AIMGet = AIMGet & vbCrLf
End Function

Function GrabCookie(theStuff, theCookie)
Dim A1, A2
A1 = Split(theStuff, vbCrLf)
For i = 0 To UBound(A1)
    If Len(A1(i)) >= 12 + Len(theCookie) Then
        If Left(A1(i), 12 + Len(theCookie)) = "Set-Cookie: " & theCookie Then
            A2 = Split(A1(i), ";")
            GrabCookie = Right(A2(0), Len(A2(0)) - 12)
            Exit Function
        End If
    End If
    DoEvents
Next i
GrabCookie = False
End Function

