VB_Name = "Utility"
Sub Scan_For_Broken_Links()
' Checking links in the sheet using http requests
If MsgBox("Is the Active Sheet a Sheet with Hyperlinks you would like to check?", vbOKCancel) = vbCancel Then
    Exit Sub
End If

Dim aLink As Hyperlink
Dim strUrl As String
Dim objhttp As Object

Dim fsoFSO As Object
Set fsoFSO = CreateObject("Scripting.FileSystemObject")

On Error Resume Next
For Each aLink In Cells.Hyperlinks
    strUrl = aLink.Address
    
    aLink.Parent.Interior.ColorIndex = xlNone
    
    If Left(strUrl, 5) <> "http:" Then
        ' if link doesn't have http at the start add it.
        strUrl = "http:" & strUrl
    End If
    Application.StatusBar = "Testing link: " & strUrl
    setobjhttp = CreateObject("MSXML2.XMLHTTP")
    objhttp.Open "HEAD", strUrl, False
    objhttp.Send
    
    If objhttp.StatusText <> "OK" Then
        aLink.Parent.Interior.Color = 255
        
        If Left(strUrl, 4) <> "http" Then
            strUrl = ActiveWorkbook.Path & "\" & aLink.Address
            ' In case the link is a file in the system (it will show up as relative path)
            If fsoFSO.FileExists(strUrl) = True Then
                aLink.Parent.Interior.ColorIndex = xlNone
            End If
        End If
    End If
Next aLink


Application.StatusBar = False
Set aLink = Nothing
Set objhttp = Nothing
Set fsoFSO = Nothing

On Error GoTo 0
MsgBox ("Checking Complete!" & vbCrLf & vbCrLf & "Cells with broken link are highlighted in red")

End Sub


End Function

Sub mergeCollection(toBeMerged As Collection, col As Collection)
' Merging two collections the first one is the merged one

For Each aItem In col
toBeMerged.Add (aItem)
Next aItem

End Sub

Function CutLeftUntil(str As String, cutTo As String, Optional include As Boolean = False, Optional timesToCut As Integer = 1)
' Cuts text from the left until certain String is reached a certain times
' By default the string will be removed
' For example if str = "192.168.1.1" and we cut to "." with include on and 2 times we get "192.168"

For i = 1 To timesToCut
    For counter = Len(str) To 1 Step -1
        If Mid(str, counter, Len(cutTo)) = cutTo Then
            If include Then
                str = Left(str, counter - 1)
            Else
                str = Left(str, counter + Len(cutTo) - 1)
            End If
            GoTo StopCutting
        End If
    Next
StopCutting:
Next

CutLeftUntil = str
End Function

Function CutRightUntil(str As String, cutTo As String, Optional include As Boolean = False, Optional timesToCut As Integer = 1)
' Cuts text from the left until certain String is reached a certain times
' By default the string will be removed
' For example if str = "192.168.1.1" and we cut to "." with include on and 2 times we get "1.1"

For i = 1 To timesToCut
    For counter = 1 To Len(str)
        If Mid(str, counter, Len(cutTo)) = cutTo Then
            If include Then
                str = Right(str, Len(str) - counter - Len(cutTo) + 1)
            Else
                str = Right(str, Len(str) - counter + 1)
            End If
            GoTo StopCutting
        End If
    Next
StopCutting:
Next

CutRightUntil = str
End Function

Sub Replace_Links()
' Used to change the url of multiple same links at once

Dim aLink As Hyperlink

oldstr = Application.InputBox("Old link: ", "Replace Hyperlinks", Type:=2)
newstr = Application.InputBox("New link: ", "Replace Hyperlinks", Type:=2)

For Each aLink In Cells.Hyperlinks
    aLink.Address = Replace(aLink.Address, oldstr, newstr)
Next aLink

End Sub

Sub Replace_Links_Selection()
' Used to change the url of multiple same links at once in a selection

Dim aLink As Hyperlink

oldstr = Application.InputBox("Old link: ", "Replace Hyperlinks", Type:=2)
newstr = Application.InputBox("New link: ", "Replace Hyperlinks", Type:=2)

For Each aLink In Selection.Hyperlinks
    aLink.Address = Replace(aLink.Address, oldstr, newstr)
Next aLink

End Sub


Sub RetPropRev()
' Prints the file's revision

On Error Resume Next
For Each prop In ActiveWorkbook.ContentTypeProperties
    If prop.Name = "תוקף/מהדורה" Or prom.Name = "מהדורה" Then
        MsgBox prop.Value
    End If
Next prop

End Sub