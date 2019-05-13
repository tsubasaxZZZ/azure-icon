Attribute VB_Name = "Module1"
Sub ImportPic(oSld, parentPath)

Dim strTemp As String
Dim strPath As String
Dim strFileSpec As String
Dim oPic As Shape

strPath = parentPath & "\"
strFileSpec = "*.svg"

strTemp = Dir(strPath & strFileSpec)

Dim l As Integer
Dim t As Integer

l = 10
t = 50
picNum = 1

'title
Set tTxt = oSld.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 300, 100)
tTxt.TextFrame.TextRange.Font.Size = 18
pos = InStrRev(parentPath, "\")
tTxt.TextFrame.TextRange.Text = Mid(parentPath, pos + 1)

Do While strTemp <> ""
    Set oPic = oSld.Shapes.AddPicture(FileName:=strPath & strTemp, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=l, _
    Top:=t, _
    Width:=-1, _
    Height:=-1)
    oPic.LockAspectRatio = msoTrue
    oPic.Height = 50

   iconName = Replace(strTemp, ".svg", "")
   Set oTxt = oSld.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t + 50, 100, 20)
   oTxt.TextFrame.TextRange.Font.Size = 10
   oTxt.TextFrame.TextRange.Text = iconName


    l = l + 120
    
    If picNum Mod 8 = 0 Then
        t = t + 100
        l = 10
    End If
    
    picNum = picNum + 1
    
    strTemp = Dir
    
Loop

End Sub

Sub loopdir()

strParentPath = InputBox("アイコンセットの入ったフォルダを指定してください。" & vbCrLf & "例) C:\temp\Microsoft_Cloud_AI_Azure_" & vbCrLf & "Service_Icon_Set_2019_05_08_2")
If strParentPath = "" Then
    MsgBox ("入力がありません。終了します。")
    Exit Sub
End If

dirPathTemp = Dir(strParentPath, vbDirectory)

'-------------------------- 親フォルダ
With CreateObject("Scripting.FileSystemObject")
    For Each f In .GetFolder(strParentPath).SubFolders
        Set oSld = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
        Call ImportPic(oSld, f.Path)
        
    Next f
End With



End Sub
