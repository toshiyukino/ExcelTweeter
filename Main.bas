Attribute VB_Name = "Main"
Option Explicit
'--------------------------------------------------------------------
'API宣言
'--------------------------------------------------------------------
Private Declare Function CreateURLMoniker Lib "urlmon.dll" _
  (ByVal pMkCtx As Long, _
  ByVal szURL As Long, _
  ByRef ppmk As Long) As Long
Private Declare Function ShowHTMLDialog Lib "mshtml.dll" _
  (ByVal hwndParent As Long, _
  ByVal pMk As Long, _
  ByVal pvarArgIn As Long, _
  ByVal pchOptions As Long, _
  ByVal pvarArgOut As Long) As Long
Private Const S_OK = 0
Private Const E_OUTOFMEMORY = &H8007000E
Private Const MK_E_SYNTAX = &H800401E4
'====================================================================

'--------------------------------------------------------------------
'設定
'--------------------------------------------------------------------
Private Const tl_count = 20       'タイムライン数
Private Const use_syntax = True   'タイムラインの色分け表示

'I/Fはシンプルにしたいので、ショートカットキーで機能を実装
'（極力左手で済むように）
'ショートカットキー
Private Const sck_post = "^t"   'ツイート           ctrl + t
Private Const sck_quot = "^q"   '旧タイプリツート　 ctrl + q
Private Const sck_reply = "^r"  '返信               ctrl + r
Private Const sck_rtwt = "^+r"  '公式リツイート     ctrl + alt + r
Private Const sck_fvadd = "^+f" 'お気に入りに登録   ctrl + alt + f
'====================================================================


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
  
    'シートダブルクリックイベント
    .Worksheets(1).OnDoubleClick = "NextTimeLine"
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
    .OnKey sck_post, "DoPost" '念のため
    .OnKey sck_quot, "DoQuottweet"
    .OnKey sck_rtwt, "DoRetweet"
    .OnKey sck_reply, "DoReply"
    .OnKey sck_fvadd, "DoFavPost"
  End With
End Sub

'ショートカットキー削除
Private Sub ResetShortcutKey()
  With Application
    .OnKey sck_quot
    .OnKey sck_rtwt
    .OnKey sck_reply
    .OnKey sck_fvadd
  End With
End Sub

'投稿
Private Sub DoPost()
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

'ツイート取得
Sub GetTweet()
  If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then
    Application.StatusBar = "ツイートを取得中..."
    Debug.Print StatusShow(ActiveSheet.Cells(ActiveCell.Row, 1).Value)
    Application.StatusBar = False
  End If
End Sub

'旧タイプリツイート
Sub DoQuottweet()
  Dim RegEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Dim msg As String
  Set RegEx = CreateObject("VBScript.RegExp")
  
  RegEx.IgnoreCase = True
  RegEx.Global = True
  'regEx.Pattern = "^.+?:\s(.+?)\d\d-\d\d-\d\d\s\d\d:\d\d$"
  RegEx.Pattern = "^(.+?)\d\d-\d\d-\d\d\s\d\d:\d\d$"
  
  If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then
  
    Set Matches = RegEx.Execute(ActiveCell.Value)
    
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
  Dim RegEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Dim msg As String
  Set RegEx = CreateObject("VBScript.RegExp")
  
  RegEx.IgnoreCase = True
  RegEx.Global = True
  RegEx.Pattern = "^.+?:\s(.+?)\d\d-\d\d-\d\d\s\d\d:\d\d$"
  
  If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then
  
    Set Matches = RegEx.Execute(ActiveCell.Value)
    
    If Matches Is Nothing Then
      MsgBox "ツイートが取得できません。" & vbCrLf & "確認してもう一度実行してください", vbCritical
      Exit Sub
    End If
    
    msg = Trim(Matches(0).SubMatches(0))
    If MsgBox("以下の内容をリツートします。" & vbCrLf & vbCrLf & _
       "「" & msg & "」" & vbCrLf & vbCrLf & _
       "よろしいですか？", vbYesNo + vbDefaultButton2) = vbYes Then
      Application.StatusBar = "ツイートを送信中..."
      Debug.Print TweetPost(msg, Re_Tweet, ActiveSheet.Cells(ActiveCell.Row, 1).Value)
      Application.StatusBar = False
    End If
  End If
End Sub

'返信
Sub DoReply()
  Dim RegEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Dim msg As String
  Set RegEx = CreateObject("VBScript.RegExp")
  
  RegEx.IgnoreCase = True
  RegEx.Global = True
  RegEx.Pattern = "^(.+?):\s"
  
  If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then
  
    Set Matches = RegEx.Execute(ActiveCell.Value)
    
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

'公式リツイート
Sub DoFavPost()
  Dim RegEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Dim msg As String
  Set RegEx = CreateObject("VBScript.RegExp")
  
  RegEx.IgnoreCase = True
  RegEx.Global = True
  RegEx.Pattern = "^.+?:\s(.+?)\d\d-\d\d-\d\d\s\d\d:\d\d$"
  
  If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then
  
    Set Matches = RegEx.Execute(ActiveCell.Value)
    
    If Matches Is Nothing Then
      MsgBox "ツイートが取得できません。" & vbCrLf & "確認してもう一度実行してください", vbCritical
      Exit Sub
    End If
    
    msg = Trim(Matches(0).SubMatches(0))
    If MsgBox("以下のツイートをお気に入りに登録します。" & vbCrLf & vbCrLf & _
       "「" & msg & "」" & vbCrLf & vbCrLf & _
       "よろしいですか？", vbYesNo + vbDefaultButton2) = vbYes Then
      Application.StatusBar = "お気に入りに登録中..."
      Debug.Print TweetPost(msg, Fv_Post, ActiveSheet.Cells(ActiveCell.Row, 1).Value)
      Application.StatusBar = False
    End If
  End If
End Sub

'ホームタイムライン
Sub TL_Home()
  PrintTimeLine home_timeline
End Sub

'自分の名前を含むタイムライン
Sub TL_Reply()
  PrintTimeLine mentions
End Sub

'自分ツイートタイムライン
Sub TL_User()
  PrintTimeLine user_timeline
End Sub

'特定の人のツイートタイムライン
Sub TL_UserTimeLine()
  Dim msg As String
  Dim dic As Object
  msg = InputBox("What is UserName ?")
  If msg <> "" Then
    Set dic = CreateObject("Scripting.Dictionary")
    dic("screen_name") = msg
    PrintTimeLine user_timeline, , dic
  End If
End Sub

'お気に入り一覧
Sub TL_Fav()
  PrintTimeLine favorites_timeline
End Sub

'続きの表示用
Sub NextTimeLine()
  Dim startRow As Long
  Dim strStatusId As String
  Dim dic As Object
  
  strStatusId = ActiveCell.Offset(-1, -1).Value
  startRow = ActiveCell.Row - 1
  
  'ステータスIDがあるか確認
  If strStatusId = "" Then  '無ければ
    startRow = 2
    strStatusId = "0"
  End If
  
  If ActiveCell.Value = "More..." Then
    Set dic = CreateObject("Scripting.Dictionary")
    dic("max_id") = strStatusId
    PrintTimeLine ActiveSheet.Range("A1").Value, startRow, dic
  Else
    SendKeys "{F2}"
  End If
End Sub

'各種タイムライン処理
Private Sub PrintTimeLine(tl_name As TimeLineName, _
                          Optional start_row As Long = 2, _
                          Optional optdic As Object = Nothing)
  Dim vntTimeLine As Variant
  Dim i As Long, j As Long
  Dim strTemp As String
  Dim vntdata() As Variant
  Dim strStatusId As String
  
  'シート初期化
  With ThisWorkbook.Worksheets(1)
    '転記行が初期値ならシートを初期化する
    If start_row = 2 Then
      .Cells.ClearContents  '転記行が指定されていなかったら削除
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
        .IndentLevel = 1  'アイコン表示用
      End With
      'セルに何のタイムラインか書いておく
      .Range("A1").Value = tl_name
      .Pictures.Delete
    End If
  End With
  
  '計算を止める
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  
  Application.StatusBar = "タイムラインを取得中..."
  vntTimeLine = GetTimeLine(tl_count, tl_name, optdic)
  
  Application.StatusBar = "タイムラインを転記中..."
  If IsArray(vntTimeLine) Then
    ReDim vntdata(0 To UBound(vntTimeLine), 0 To 1)
    For i = 0 To UBound(vntTimeLine)
        If Trim(vntTimeLine(i, 1)) <> "" Then
          DoEvents
          vntdata(i, 0) = "'" & vntTimeLine(i, 0) '桁が大きいので文字として記入
          vntdata(i, 1) = vntTimeLine(i, 1) & ": " & vntTimeLine(i, 2) & " " & vntTimeLine(i, 3)
        End If
    Next
    
    With ThisWorkbook.Worksheets(1)
      .Cells(start_row, 1).Resize(UBound(vntTimeLine) + 1, 2).Value = vntdata
      
      'シンタックスハイライト(たぶん遅くなる)
      j = start_row
      Do Until .Cells(j, 1).Value = ""
        Syntax .Cells(j, 2)
        j = j + 1
      Loop
      
      'アイコン表示
      For i = 0 To UBound(vntTimeLine)
        AddProfileImage .Cells(i + start_row, 2), vntTimeLine(i, 4)
      Next
      
      '追加取得用
      .Cells(.Cells(.Cells.Rows.Count, 2).End(xlUp).Row + 1, 2).Value = "More..."
    End With
  End If
  
  Application.StatusBar = False
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
  Application.EnableEvents = True
End Sub

'見やすく
Sub Syntax(Target As Range)
  Dim strText As String
  Dim RegEx As Object 'VBScript_RegExp_55.RegExp
  Dim Match As Object ' VBScript_RegExp_55.Match
  Dim Matches As Object 'VBScript_RegExp_55.MatchCollection
  Set RegEx = CreateObject("VBScript.RegExp")
  
  RegEx.IgnoreCase = True
  RegEx.Global = True
  
  strText = Target.Value
  
  'user
  RegEx.Pattern = "^.+?:"
  Set Matches = RegEx.Execute(strText)
  For Each Match In Matches
    With Target.Characters(Match.FirstIndex + 1, Match.Length).Font
      .ColorIndex = 54
    End With
  Next
    
  'create time
  'regEx.Pattern = "\s\s.+$"
  RegEx.Pattern = "\d\d-\d\d-\d\d\s\d\d:\d\d$"
  Set Matches = RegEx.Execute(strText)
  For Each Match In Matches
    With Target.Characters(Match.FirstIndex + 1, Match.Length).Font
      .ColorIndex = 10
    End With
  Next
    
  'url
  RegEx.Pattern = "(http|https|ftp)://\S+"
  Set Matches = RegEx.Execute(strText)
  For Each Match In Matches
    With Target.Characters(Match.FirstIndex + 1, Match.Length).Font
      .ColorIndex = 5
      .Underline = xlUnderlineStyleSingle
    End With
  Next
  
  'mail
  RegEx.Pattern = "\w+@\w+"
  Set Matches = RegEx.Execute(strText)
  For Each Match In Matches
    With Target.Characters(Match.FirstIndex + 1, Match.Length).Font
      .ColorIndex = 5
      .Underline = xlUnderlineStyleSingle
    End With
  Next
  
  'hashタグ
  RegEx.Pattern = "#\w+"
  Set Matches = RegEx.Execute(strText)
  For Each Match In Matches
    Target.Characters(Match.FirstIndex + 1, Match.Length).Font.ColorIndex = 48
  Next
  
  'reply
  RegEx.Pattern = "@\w+"
  Set Matches = RegEx.Execute(strText)
  For Each Match In Matches
    Target.Characters(Match.FirstIndex + 1, Match.Length).Font.ColorIndex = 5
  Next

End Sub

'イメージ表示用
Sub AddProfileImage(strTarget As Range, ByVal strPath As String)
  Dim myShape As Shape
  
  If strPath = "" Then Exit Sub
  
  Set myShape = ThisWorkbook.Worksheets(1).Shapes.AddPicture( _
        Filename:=strPath, _
        LinkToFile:=True, _
        SaveWithDocument:=False, _
        Left:=strTarget.Left, _
        Top:=strTarget.Top, _
        Width:=16, _
        Height:=16)
End Sub



