Attribute VB_Name = "Main"
Option Explicit

Private Const tl_count = 50
'�V���[�g�J�b�g�L�[
Private Const sck_post = "^t"  '�c�C�[�g
Private Const sck_quot = "^q"  '���^�C�v���c�[�g
Private Const sck_rtwt = "^+r"  '�������c�C�[�g
Private Const sck_reply = "^r" '�ԐM

'�u�b�N�I�[�v�����Ɏ��s
Private Sub Auto_Open()
  On Error Resume Next
  '�V���[�g�J�b�g�L�[
  With Application
    .OnKey sck_post, "DoPost" '�|�X�g�͑S�Ẵu�b�N���ɗL���ɂ���B
    
    '�^�C�����C���\���V�[�g���A�N�e�B�u�ȂƂ��Ɏ��s
    Call SetShortcutKey
    .Worksheets(1).OnSheetActivate = "SetShortcutKey"
    .Worksheets(1).OnSheetDeactivate = "ResetShortcutKey"
  End With
End Sub

'�u�b�N�I�[�v�����Ɏ��s
Private Sub Auto_Close()
  On Error Resume Next
  '�V���[�g�J�b�g�L�[
  With Application
    .OnKey sck_post
    Call ResetShortcutKey
  End With
End Sub

Private Sub SetShortcutKey()
  With Application
    .OnKey sck_quot, "DoQuottweet"  '���c�[�g(���^�C�v)
    .OnKey sck_rtwt, "DoRetweet"
  End With
End Sub

Private Sub ResetShortcutKey()
  With Application
    .OnKey sck_quot
    .OnKey sck_rtwt
  End With
End Sub

Sub DoPost()
Attribute DoPost.VB_Description = "�Ԃ₫���|�X�g���܂�"
Attribute DoPost.VB_ProcData.VB_Invoke_Func = " \n14"
  Dim msg As String
  msg = InputBox("what are you doing?")
  If msg <> "" And Len(msg) < 141 Then
    Application.StatusBar = "�c�C�[�g�𑗐M��..."
    If MsgBox("tweet ok?", vbYesNo) = vbYes Then
      Debug.Print TweetPost(msg)
    End If
    Application.StatusBar = False
  End If
End Sub

Sub DoQuottweet()
Attribute DoQuottweet.VB_Description = "���^�C�v�̃��c�[�g�����܂��B"
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
      MsgBox "�c�C�[�g���擾�ł��܂���B" & vbCrLf & "�蓮�ōs����������x���s���Ă�������", vbCritical
      Exit Sub
    End If
    
    msg = "QT @" & Trim(Matches(0).SubMatches(0))
    msg = InputBox("what are you doing?", , msg)
    If msg <> "" And Len(msg) < 141 Then
      If MsgBox("tweet ok?", vbYesNo) = vbYes Then
        Application.StatusBar = "�c�C�[�g�𑗐M��..."
        Debug.Print TweetPost(msg)
        Application.StatusBar = False
      End If
    End If
  End If
End Sub

Sub DoRetweet()
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
      MsgBox "�c�C�[�g���擾�ł��܂���B" & vbCrLf & "�m�F���Ă�����x���s���Ă�������", vbCritical
      Exit Sub
    End If
    
    msg = Trim(Matches(0).SubMatches(0))
    If MsgBox("�ȉ��̓��e�����c�[�g���܂��B" & vbCrLf & vbCrLf & _
       "�u" & msg & "�v" & vbCrLf & vbCrLf & _
       "��낵���ł����H", vbYesNo + vbDefaultButton2) = vbYes Then
      If MsgBox("tweet ok?", vbYesNo) = vbYes Then
        Application.StatusBar = "�c�C�[�g�𑗐M��..."
        Debug.Print TweetPost(msg, Re_Tweet, ActiveSheet.Cells(ActiveCell.Row, 1).Value)
        Application.StatusBar = False
      End If
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
  
  '�N���A
  ThisWorkbook.Worksheets(1).Cells.ClearContents
  
  '�v�Z���~�߂�
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.StatusBar = "�^�C�����C�����擾��..."
  
  vntTimeLine = GetTimeLine(tl_count, tl_name)
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
          Syntax .Cells(j, 2) '���Ԃ�x���Ȃ�
        End If
      End With
      j = j + 1
    Next
  End If
  
  Application.StatusBar = False
  Application.ScreenUpdating = True
  Application.Calculation = xlCalculationAutomatic
End Sub

'���₷��
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
  
  'hash�^�O
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


