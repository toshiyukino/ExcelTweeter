Attribute VB_Name = "Main"
Option Explicit

Private Const tl_count = 50
'ショートカットキー
Private Const sck_post = "^t"   'ツイート           ctrl + t
Private Const sck_quot = "^q"   '旧タイプリツート　 ctrl + q
Private Const sck_reply = "^r"  '返信               ctrl + r
Private Const sck_rtwt = "^+r"  '公式リツイート     ctrl + alt + r

'ブックオープン時に実行
Private Sub Auto_Open()
  On Error Resume Next
  'ショートカットキー
  With Application
    '全ブックどこでも
    .OnKey sck_post, "DoPost" 'ポストは全てのブック時に有効にする。
    
    'タイムライン表示シートがアクティブなときに実行
    Call SetShortcutKey
    .Worksheets(1).OnSheetActivate = "SetShortcutKey"
    
    '非アクティブ時
    .Worksheets(1).OnSheetDeactivate = "ResetShortcutKey"
  End With
End Sub

'ブッククローズ時に実行
Private Sub Auto_Close()
  On Error Resume Next
  'ショートカットキー
  With Application
    .OnKey sck_post
    Call ResetShortcutKey
  End With
End Sub

'ショートカットキー追加
Private Sub SetShortcutKey()
  With Application
    .OnKey sck_quot, "DoQuottweet"  'リツート(旧タイプ)
    .OnKey sck_rtwt, "DoRetweet"
    .OnKey sck_reply, "DoReply"
  End With
End Sub

'ショートカットキー削除
Private Sub ResetShortcutKey()
  With Application
    .OnKey sck_quot
    .OnKey sck_rtwt
    .OnKey sck_reply
  End With
End Sub

'投稿
Sub DoPost()
Attribute DoPost.VB_Description = "つぶやきをポストします"
Attribute DoPost.VB_ProcData.VB_Invoke_Func = " \n14"
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

'旧タイプリツイート
Sub DoQuottweet()
Attribute DoQuottweet.VB_Description = "旧タイプのリツートをします。"
Attribute DoQuottweet.VB_ProcData.VB_Invoke_Func = " \n14"
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
    
    If Matches Is Nothing Then
      MsgBox "ツイートが取得できません。" & vbCrLf & "手動で行うかもう一度実行してください", vbCritical
      Exit Sub
    End If
    
    msg = " QT @" & Trim(Matches(0).SubMatches(0)) & " "
    Application.SendKeys "{HOME}", False
    msg = InputBox("what are you doing?", , msg)
    If msg <> "" And Len(msg) < 141 Then
      If MsgBox("tweet ok?", vbYesNo) = vbYes Then
        Application.StatusBar = "ツイートを送信中..."
        Debug.Print TweetPost(msg, Qt_Tweet, ActiveSheet.Cells(ActiveCell.Row, 1).Value)
        Application.StatusBar = False
      End If
    End If
  End If
End Sub

'公式リツイート
Sub DoRetweet()
Attribute DoRetweet.VB_Description = "公式リツイートをします。"
Attribute DoRetweet.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim regEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Dim msg As String
  Set regEx = CreateObject("VBScript.RegExp")
  
  regEx.IgnoreCase = True
  regEx.Global = True
  regEx.Pattern = "^.+?:\s(.+?)\d\d-\d\d-\d\d\s\d\d:\d\d$"
  
  If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then
  
    Set Matches = regEx.Execute(ActiveCell.Value)
    
    If Matches Is Nothing Then
      MsgBox "ツイートが取得できません。" & vbCrLf & "確認してもう一度実行してください", vbCritical
      Exit Sub
    End If
    
    msg = Trim(Matches(0).SubMatches(0))
    If MsgBox("以下の内容をリツートします。" & vbCrLf & vbCrLf & _
       "「" & msg & "」" & vbCrLf & vbCrLf & _
       "よろしいですか？", vbYesNo + vbDefaultButton2) = vbYes Then
      If MsgBox("tweet ok?", vbYesNo) = vbYes Then
        Application.StatusBar = "ツイートを送信中..."
        Debug.Print TweetPost(msg, Re_Tweet, ActiveSheet.Cells(ActiveCell.Row, 1).Value)
        Application.StatusBar = False
      End If
    End If
  End If
End Sub

'返信
Sub DoReply()
Attribute DoReply.VB_Description = "返信します。"
Attribute DoReply.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim regEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Dim msg As String
  Set regEx = CreateObject("VBScript.RegExp")
  
  regEx.IgnoreCase = True
  regEx.Global = True
  regEx.Pattern = "^(.+?):\s"
  
  If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then
  
    Set Matches = regEx.Execute(ActiveCell.Value)
    
    If Matches Is Nothing Then
      MsgBox "ツイートが取得できません。" & vbCrLf & "確認してもう一度実行してください", vbCritical
      Exit Sub
    End If
    
    Application.SendKeys "{RIGHT}", False
    msg = InputBox("Re: " & ActiveSheet.Cells(ActiveCell.Row, 2).Value, , "@" & Trim(Matches(0).SubMatches(0)) & " ")
    If msg <> "" And Len(msg) < 141 Then
      If MsgBox("tweet ok?", vbYesNo) = vbYes Then
        Application.StatusBar = "ツイートを送信中..."
        Debug.Print TweetPost(msg, Reply_Tweet, ActiveSheet.Cells(ActiveCell.Row, 1).Value)
        Application.StatusBar = False
      End If
    End If
  End If
End Sub

'ホームタイムライン
Sub TL_Home()
Attribute TL_Home.VB_Description = "タイムラインを表示"
Attribute TL_Home.VB_ProcData.VB_Invoke_Func = " \n14"
  PrintTimeLine home_timeline
End Sub

'自分の名前を含むタイムライン
Sub TL_Reply()
Attribute TL_Reply.VB_Description = "自分の名前を含むタイムラインを表示"
Attribute TL_Reply.VB_ProcData.VB_Invoke_Func = " \n14"
  PrintTimeLine mentions
End Sub

'各種タイムライン処理
Private Sub PrintTimeLine(tl_name As TimeLineName)
  Dim vntTimeLine As Variant
  Dim i As Long, j As Long
  Dim strTemp As String
  
  'シート初期化
  With ThisWorkbook.Worksheets(1)
    .Cells.ClearContents
    With .Columns(1) 'ステータスｉｄ表示用
      '.WrapText = False
      .Hidden = True
    End With
    With .Columns(2) 'タイムライン表示用
      .ColumnWidth = 100
      .VerticalAlignment = xlVAlignTop
      .WrapText = True
      .Font.ColorIndex = xlAutomatic
      .Font.Size = 9
    End With
  End With
  
  '計算を止める
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.StatusBar = "タイムラインを取得中..."
  
  vntTimeLine = GetTimeLine(tl_count, tl_name)
  If IsArray(vntTimeLine) Then
    j = 2
    For i = 0 To UBound(vntTimeLine)
      With ThisWorkbook.Worksheets(1)
        If Trim(vntTimeLine(i, 1)) <> "" Then
          .Cells(j, 1).Value = "'" & vntTimeLine(i, 0)  '桁が大きいので文字として記入
          strTemp = vntTimeLine(i, 1) & ": " & vntTimeLine(i, 2) & " " & vntTimeLine(i, 3)
          .Cells(j, 2).Value = strTemp
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


