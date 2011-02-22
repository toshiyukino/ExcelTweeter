Attribute VB_Name = "Main"
Option Explicit

Private Const tl_count = 50

Sub DoPost()
Attribute DoPost.VB_Description = "つぶやきをポストします"
Attribute DoPost.VB_ProcData.VB_Invoke_Func = "t\n14"
  Dim msg As String
  msg = InputBox("what are you doing?")
  If msg <> "" And Len(msg) < 141 Then
    Application.StatusBar = "ツイートを送信中..."
    If MsgBox("tweet ok?", vbYesNo) = vbYes Then
      Debug.Print TweetPost(msg)
    End If
    Application.StatusBar = False
  End If
End Sub

Sub DoRetweet()
Attribute DoRetweet.VB_Description = "旧タイプのリツートをします。"
Attribute DoRetweet.VB_ProcData.VB_Invoke_Func = "r\n14"
  Dim regEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Dim msg As String
  Set regEx = CreateObject("VBScript.RegExp")
  
  regEx.IgnoreCase = True
  regEx.Global = True
  'regEx.Pattern = "^.+?:\s(.+?)\d\d-\d\d-\d\d\s\d\d:\d\d$"
  regEx.Pattern = "^(.+?)\d\d-\d\d-\d\d\s\d\d:\d\d$"
  
  If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then
  
    Set Matches = regEx.Execute(ActiveCell.Value)
    msg = "RT @" & Trim(Matches(0).SubMatches(0))
    msg = InputBox("what are you doing?", , msg)
    If msg <> "" And Len(msg) < 141 Then
      Application.StatusBar = "ツイートを送信中..."
      If MsgBox("tweet ok?", vbYesNo) = vbYes Then
        Debug.Print TweetPost(msg)
      End If
      Application.StatusBar = False
    End If
  End If
End Sub

Sub TL_Home()
  PrintTimeLine home_timeline
End Sub

Sub TL_Reply()
  PrintTimeLine mentions
End Sub

Private Sub PrintTimeLine(tl_name As TimeLineName)
  Dim vntTimeLine As Variant
  Dim i As Long, j As Long
  Dim strTemp As String
  
  '計算を止める
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.StatusBar = "タイムラインを取得中..."
  
  vntTimeLine = GetTimeLine(tl_count, tl_name)
  ThisWorkbook.Worksheets(1).Cells.ClearContents
  If IsArray(vntTimeLine) Then
    j = 2
    For i = 0 To UBound(vntTimeLine)
      With ThisWorkbook.Worksheets(1)
        If Trim(vntTimeLine(i, 1)) <> "" Then
          .Cells(j, 1).Value = vntTimeLine(i, 0)
          strTemp = vntTimeLine(i, 1) & ": " & vntTimeLine(i, 2) & " " & vntTimeLine(i, 3)
          .Cells(j, 2).Value = strTemp
          .Cells(j, 2).VerticalAlignment = xlVAlignTop
          .Cells(j, 2).WrapText = True
          .Cells(j, 2).Font.ColorIndex = xlAutomatic
          Syntax .Cells(j, 2) 'たぶん遅くなる
        End If
      End With
      j = j + 1
    Next
  End If
  
  Application.StatusBar = False
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
End Sub

'見やすく
Sub Syntax(Target As Range)
  Dim regEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Set regEx = CreateObject("VBScript.RegExp")
  
  regEx.IgnoreCase = True
  regEx.Global = True
  
  'user
  regEx.Pattern = "^.+?:"
  Set Matches = regEx.Execute(Target.Value)
  For Each Match In Matches
    With Target.Characters(Match.FirstIndex + 1, Match.Length).Font
      .ColorIndex = 54
    End With
  Next
    
  'create time
  'regEx.Pattern = "\s\s.+$"
  regEx.Pattern = "\d\d-\d\d-\d\d\s\d\d:\d\d$"
  Set Matches = regEx.Execute(Target.Value)
  For Each Match In Matches
    With Target.Characters(Match.FirstIndex + 1, Match.Length).Font
      .ColorIndex = 10
    End With
  Next
    
  'url
  regEx.Pattern = "(http|https|ftp)://\S+"
  Set Matches = regEx.Execute(Target.Value)
  For Each Match In Matches
    With Target.Characters(Match.FirstIndex + 1, Match.Length).Font
      .ColorIndex = 5
      .Underline = xlUnderlineStyleSingle
    End With
  Next
  
  'mail
  regEx.Pattern = "\w+@\w+"
  Set Matches = regEx.Execute(Target.Value)
  For Each Match In Matches
    With Target.Characters(Match.FirstIndex + 1, Match.Length).Font
      .ColorIndex = 5
      .Underline = xlUnderlineStyleSingle
    End With
  Next
  
  'hashタグ
  regEx.Pattern = "#\w+"
  Set Matches = regEx.Execute(Target.Value)
  For Each Match In Matches
    Target.Characters(Match.FirstIndex + 1, Match.Length).Font.ColorIndex = 48
  Next
  
  'reply
  regEx.Pattern = "@\w+"
  regEx.IgnoreCase = True
  regEx.Global = True
  Set Matches = regEx.Execute(Target.Value)
  For Each Match In Matches
    Target.Characters(Match.FirstIndex + 1, Match.Length).Font.ColorIndex = 5
  Next

End Sub


