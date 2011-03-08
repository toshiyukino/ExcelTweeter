Attribute VB_Name = "ExcelTweeter"
Option Explicit


'-----------------------------------
'UserAgent
'-----------------------------------
Private Const UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.1) "
'-----------------------------------
'PROXY 設定(使う場合はユーザ名：パスワードで指定)
'-----------------------------------
Private Const proxy_user = ""
Private Const proxy_server = ""
'-----------------------------------
'TwitterAPI URL
'-----------------------------------
Private Const reqt_url = "http://api.twitter.com/oauth/request_token"
Private Const auth_url = "https://api.twitter.com/oauth/authorize"
Private Const acct_url = "http://api.twitter.com/oauth/access_token"
Private Const post_url = "https://api.twitter.com/1/statuses/update.xml"
Private Const retw_url = "https://api.twitter.com/1/statuses/retweet/" ' & statusid & ".xml"
'Private Const frtl_url = "https://api.twitter.com/1/statuses/friends_timeline.xml"
'Private Const hmtl_url = "https://api.twitter.com/1/statuses/home_timeline.xml"
Private Const timeline_url = "https://api.twitter.com/1/statuses/"
Public Enum TimeLineName
  home_timeline = 1
  friends_timeline = 2
  user_timeline = 3
  replies = 4
  mentions = 5
  retweeted_by_me = 6
  retweeted_to_me = 7
  retweets_of_me = 8
End Enum
Public Enum TweetType
  Default_Tweet = 1 '普通のポスト
  Reply_Tweet = 2   '返信
  Re_Tweet = 3      '公式リツート
  Rt_Tweet = 4      '非公式リツイート
  Qt_Tweet = 5      '引用ツイート
End Enum
'-----------------------------------
'ConsumerKey
'-----------------------------------
Private Const Consumer_key = "TJ0pecuWf8ctAExNQbKLQ"
Private Const Consumer_secret = "oZkjqJCb0nIdwK3RmPe2lEcsw9QzK8Q0NQLolqDqMwY"
'-----------------------------------
'API 宣言
'-----------------------------------
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" _
  (ByRef phProv As Long, ByVal pszContainer As String, _
  ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" _
  (ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" _
  (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, _
  ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
  (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" _
  (ByVal hHash As Long, ByRef pbData As Any, _
  ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
  (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Any, _
  ByRef pdwDataLen As Long, ByVal dwFlags As Integer) As Long
Private Declare Function CryptBinaryToString Lib "crypt32.dll" Alias "CryptBinaryToStringA" _
    (ByRef pbBinary As Any, ByVal cbBinary As Long, _
     ByVal dwFlags As Long, ByVal pszString As String, _
     ByRef pcchString As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'アクセストークンを保存しているファイルを削除
Public Function DelTokenFile() As Boolean
  Dim intFileNo As Integer
  Dim strFileName As String
  
  On Error GoTo ErrorHandler
  strFileName = GetTokenFileName()
  If strFileName = "" Then
    DelTokenFile = False
    Exit Function
  End If

  Kill strFileName
  DelTokenFile = True
  Exit Function

ErrorHandler:
  DelTokenFile = False
End Function

Public Function TweetPost( _
  strPost As String, _
  Optional Tweet_type As TweetType = Default_Tweet, _
  Optional strStatusID As String = "" _
) As String
  Dim xhr As Object 'MSXML2.ServerXMLHTTP60
  Dim param As Object  'Scripting.Dictionary
  Dim reqdata As String
  Dim strReqURL As String
  Dim digest As String
  Dim buf() As Byte
  Dim res As String
  Dim XMLDOM As Object 'MSXML2.DOMDocument
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim i As Long

  If Not IsArray(GetToken) Then
    If isOAuth = False Then
      MsgBox "アクセストークンが取得できませんでした", vbCritical
      TweetPost = ""
      Exit Function
    End If
  End If
  
  atoken = GetToken(0)
  atoken_secret = GetToken(1)
  
  Set param = CreateObject("Scripting.Dictionary")
  param("oauth_token") = atoken
  param("source") = "ExcelTweet"
  If Tweet_type = Re_Tweet Then
      '公式リツイート
      strReqURL = retw_url & strStatusID & ".xml"
      If strStatusID = "" Then
        MsgBox "リツート元のステータスＩＤが取得できませんでした。", vbCritical
        TweetPost = ""
        Exit Function
      End If
      param("id") = strStatusID
  Else
    strReqURL = post_url
    param("status") = UrlEncode(Left(strPost, 140)) '140文字
    '通常のポスト以外はは返信元ＩＤを入れる
    If Tweet_type <> Default_Tweet And strStatusID <> "" Then
      param("in_reply_to_status_id") = strStatusID
    End If
  End If
  strSig = MakeSignature("POST", strReqURL, param, UrlEncode(Consumer_secret) & "&" & UrlEncode(atoken_secret))
  param("oauth_signature") = UrlEncode(strSig)
  
  Set xhr = CreateRequest("POST", strReqURL)
  If xhr Is Nothing Then
    MsgBox "リクエストオブジェクトが作成できませんでした", vbCritical
    TweetPost = 0
    Exit Function
  End If
  xhr.send UrlParse(param)
  
  '読み込みが完了するまでループ
  'Do Until xhr.readyState = 4
  '  DoEvents
  '  Sleep 100
  'Loop
  
  If xhr.Status <> 200 Then
    MsgBox xhr.getAllResponseHeaders
    Set XMLDOM = xhr.responseXML
    If XMLDOM Is Nothing Then
      'MsgBox xhr.ResponseText
      TweetPost = "error"
    Else
      TweetPost = XMLDOM.selectSingleNode("hash/error").FirstChild.nodeValue
    End If
    Exit Function
  Else
    TweetPost = "ok"
  End If
End Function

'(x,0):id
'(x,1):screen name
'(x,2):tweet text
'(x,3):create time
'引数 timeline_count:取得するタイムラインの数
'     timeline_name:タイムラインの種類
'戻り値が配列か確認して使う
Public Function GetTimeLine _
  (Optional timeline_count As Long = 20, _
  Optional timeline_name As TimeLineName = home_timeline _
) As Variant
  Dim xhr As Object 'MSXML2.ServerXMLHTTP60
  Dim param As Object  'Scripting.Dictionary
  Dim reqdata As String
  Dim res As String
  Dim XMLDOM As Object 'MSXML2.DOMDocument
  Dim Statuses As Object 'MSXML2.IXMLDOMNode
  Dim objStatus As Object 'MSXML2.IXMLDOMElement
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim i As Long
  Dim strTimeLine() As String
  Dim strTemp As String
  Dim strTL_url As String

  On Error GoTo ErrorHandler
  
  If Not IsArray(GetToken) Then
    If isOAuth = False Then
      MsgBox "アクセストークンが取得できませんでした", vbCritical
      GetTimeLine = 0
      Exit Function
    End If
  End If
  
  Select Case timeline_name
    Case friends_timeline
      strTL_url = timeline_url & "friends_timeline.xml"
    Case home_timeline
      strTL_url = timeline_url & "home_timeline.xml"
    Case user_timeline
      strTL_url = timeline_url & "user_timeline.xml"
    Case replies
      strTL_url = timeline_url & "replies.xml"
    Case mentions
      strTL_url = timeline_url & "mentions.xml"
    Case retweeted_by_me
      strTL_url = timeline_url & "retweeted_by_me.xml"
    Case retweeted_to_me
      strTL_url = timeline_url & "retweeted_to_me.xml"
    Case retweets_of_me
      strTL_url = timeline_url & "retweets_of_me.xml"
    Case Else
      MsgBox "タイムラインの種類を特定できません", vbCritical
      GetTimeLine = 0
      Exit Function
  End Select
  
  atoken = GetToken(0)
  atoken_secret = GetToken(1)

  Set param = CreateObject("Scripting.Dictionary")
  param("oauth_token") = atoken 'Access_Token
  param("count") = CStr(timeline_count)
  strSig = MakeSignature("GET", strTL_url, param, UrlEncode(Consumer_secret) & "&" & UrlEncode(atoken_secret))
  param("oauth_signature") = UrlEncode(strSig)
  
  Set xhr = CreateRequest("GET", strTL_url & "?" & UrlParse(param))
  If xhr Is Nothing Then
    MsgBox "リクエストオブジェクトが作成できませんでした", vbCritical
    GetTimeLine = 0
    Exit Function
  End If
  xhr.send
  
  '読み込みが完了するまでループ
  'Do Until xhr.readyState = 4
  '  DoEvents
  '  Sleep 100
  'Loop
 
  If xhr.Status <> 200 Then
    MsgBox xhr.getAllResponseHeaders
    Set XMLDOM = xhr.responseXML
    If XMLDOM Is Nothing Then
      MsgBox xhr.responseText
    Else
      MsgBox XMLDOM.selectSingleNode("hash/error").FirstChild.nodeValue
    End If
    GetTimeLine = 0
    Exit Function
  End If
  
  ReDim strTimeLine(param("count") - 1, 3)
  Set XMLDOM = xhr.responseXML
  XMLDOM.setProperty "SelectionLanguage", "XPath"
  Set Statuses = XMLDOM.childNodes(1)
  If Statuses Is Nothing Then
    GetTimeLine = 0
    Exit Function
  End If
  
  i = 0
  For Each objStatus In Statuses.childNodes
    strTimeLine(i, 0) = CStr(objStatus.selectSingleNode("id").FirstChild.nodeValue)
    strTimeLine(i, 1) = html_escape(objStatus.selectSingleNode("user/screen_name").FirstChild.nodeValue)
    If objStatus.selectSingleNode("retweeted_status/text") Is Nothing Then
      strTimeLine(i, 2) = html_escape(objStatus.selectSingleNode("text").FirstChild.nodeValue)
    Else
      '公式リツート対応
      strTimeLine(i, 2) = objStatus.selectSingleNode("retweeted_status/user/screen_name").FirstChild.nodeValue
      strTimeLine(i, 2) = "RT @" & strTimeLine(i, 2) & ": "
      strTimeLine(i, 2) = html_escape(strTimeLine(i, 2)) & _
                          html_escape(objStatus.selectSingleNode("retweeted_status/text").FirstChild.nodeValue)
    End If
    strTimeLine(i, 3) = html_escape(Format(ConvertCreateTime(objStatus.selectSingleNode("created_at").FirstChild.nodeValue), "yy-mm-dd hh:mm"))
    i = i + 1
  Next
  GetTimeLine = strTimeLine
  Exit Function
ErrorHandler:
  GetTimeLine = 0
End Function

'コンシューマキーを使ってアクセストークンを取得してファイルに保存する。
Private Function isOAuth() As Boolean
  Dim xhr As Object 'MSXML2.ServerXMLHTTP60
  Dim param As Scripting.Dictionary
  Dim reqdata As String
  Dim res As String
  Dim otoken As String, otoken_secret As String
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim strPin As String
  Dim i As Long
  
  Set param = CreateObject("Scripting.Dictionary")

  strSig = MakeSignature("GET", reqt_url, param, UrlEncode(Consumer_secret) & "&")
  param("oauth_signature") = UrlEncode(strSig)

  Set xhr = CreateRequest("GET", reqt_url & "?" & UrlParse(param))
  If xhr Is Nothing Then
    MsgBox "リクエストオブジェクトが作成できませんでした", vbCritical
    isOAuth = False
    Exit Function
  End If
  xhr.send
  
  '取得失敗
  If xhr.Status <> 200 Then
    isOAuth = False
    Exit Function
  End If
  
  'レスポンステキスト
  res = xhr.responseText
  
  'authトークン（一時的に使う）
  otoken = GetOAuthToken(res)
  otoken_secret = GetOAuthToken_secret(res)

  'レスポンスにトークンが含まれていない場合
  If otoken = "" Or otoken_secret = "" Then
    isOAuth = False
    Exit Function
  End If

  'PIN取得の為ブラウザを起動（引数にauthトークンを指定）
  Call ShellExecute(0, "open", auth_url & "?oauth_token=" & otoken, vbNullString, vbNullString, 3)
  strPin = InputBox("pinを入力")
  If strPin = "" Then
    isOAuth = False
    Exit Function
  End If

  param.RemoveAll
  param("oauth_verifier") = strPin '今回だけ（PINコード）
  param("oauth_token") = otoken '今回だけ（authトークン）
  strSig = MakeSignature("GET", acct_url, param, UrlEncode(Consumer_secret) & "&")
  param("oauth_signature") = UrlEncode(strSig)

  Set xhr = CreateRequest("GET", acct_url & "?" & UrlParse(param))
  xhr.send
  
  If xhr.Status <> 200 Then
    isOAuth = False
    Exit Function
  End If

  'レスポンステキスト
  res = xhr.responseText
  
  atoken = GetOAuthToken(res)
  atoken_secret = GetOAuthToken_secret(res)

  'レスポンスにトークンが含まれていない場合
  If atoken = "" Or atoken_secret = "" Then
    isOAuth = False
    Exit Function
  End If

  'ファイルに保存
  If SaveToken(atoken, atoken_secret) Then
    isOAuth = True  '保存成功
  Else
    isOAuth = False '保存エラー
  End If
End Function


'xmlHttpRequestオブジェクトをオープンして返す
Private Function CreateRequest(strMethod As String, strRequestParm As String) As Object 'MSXML2.ServerXMLHTTP60
  Dim xhr As Object 'MSXML2.ServerXMLHTTP60
  Dim ua As String
  
  On Error GoTo ErrorHandler
  
  Set xhr = CreateObject("Msxml2.ServerXMLHTTP.6.0") 'MSXML2.ServerXMLHTTP60

  'プロキシ設定
  If Not proxy_server = "" Then
    xhr.SetProxy 2, proxy_server  'SXH_PROXY_SET_PROXY=2
  End If
  
  'UA設定
  If UserAgent = "" Then
    ua = "Mozilla/4.0"
  Else
    ua = UserAgent
  End If
  
  'オブジェクトを開く
  Call xhr.Open(strMethod, strRequestParm, False) '同期処理
  
  'プロキシ認証
  If Len(proxy_user) > 0 Then
    If InStr(1, proxy_user, ":") > 0 Then
      xhr.setProxyCredentials Split(proxy_user, ":")(0), Split(proxy_user, ":")(1)
    End If
  End If
  
  'リクエストヘッダーのセット
  Call xhr.setRequestHeader("User-Agent", ua)
  Call xhr.setRequestHeader("Pragma", "no-cache")
  Call xhr.setRequestHeader("Cache-Control", "Private")
  Call xhr.setRequestHeader("Expires", "-1")
  Set CreateRequest = xhr
  Exit Function
ErrorHandler:
  Set CreateRequest = Nothing
End Function

Private Function MakeSignature(strMethod As String, strUrl As String, ByRef DictionaryObject As Object, strHmacKey As String) As String
  Dim strReqData As String
  Dim buf() As Byte
  Dim strDigest As String
  Dim i As Long
  Randomize '乱数ジェネレータを初期化
  DictionaryObject("oauth_consumer_key") = Consumer_key
  DictionaryObject("oauth_signature_method") = "HMAC-SHA1"
  DictionaryObject("oauth_version") = "1.0"
  DictionaryObject("oauth_timestamp") = CStr(DateDiff("s", #1/1/1970#, Now()))
  DictionaryObject("oauth_nonce") = CStr(Int((100000000000# - 10000000 + 1) * Rnd + 10000000)) '適当に一意な値
  
  strReqData = strMethod & "&" & UrlEncode(strUrl) & "&" & UrlEncode(UrlParse(DictionaryObject))
  strDigest = hmac(strHmacKey, strReqData)
  buf = StrToBynary(strDigest)
  MakeSignature = Trim(EncodeBase64(buf)) '& vbLf
  
End Function

'wsh機能を使う(JScript)
Private Function UrlEncode(strTarget As String) As String
  Dim obj As Object
  Dim s As String
  If Len(strTarget) = 0 Then Exit Function
  Set obj = CreateObject("ScriptControl")
  obj.Language = "JScript"
  s = obj.CodeObject.encodeURIComponent(strTarget)
  '半角かっこはエンコードされないので、対策
  s = Replace(s, "(", "%28")
  s = Replace(s, ")", "%29")
  UrlEncode = s
End Function

'win32API(恐らくwin2000から動く)
Private Function EncodeBase64(bytTarget() As Byte) As String
  Dim strBase64 As String
  Dim lngBase64_Len As Long
  Dim ret As Long
  Const CRYPT_STRING_BASE64 As Long = 1
  '必要な容量を計算
  ret = CryptBinaryToString(bytTarget(0), UBound(bytTarget) + 1, CRYPT_STRING_BASE64, vbNullString, lngBase64_Len)
  If ret Then
      strBase64 = Space(lngBase64_Len)
      ret = CryptBinaryToString(bytTarget(0), UBound(bytTarget) + 1, CRYPT_STRING_BASE64, strBase64, Len(strBase64))
  End If
  EncodeBase64 = Mid(strBase64, 1, lngBase64_Len - 3)
End Function

'keyをソートして配列を返す
Private Function KeySort(dictionary_object As Object) As Variant
  Dim i As Long, j As Long
  Dim varTemp As Variant
  Dim varData As Variant
  
  If dictionary_object Is Nothing And dictionary_object.Count = 0 Then
    KeySort = 0
    Exit Function
  End If
  
  varData = dictionary_object.Keys
  
  '総当りでソート（バブルソート）
  For i = 0 To dictionary_object.Count - 1
    For j = i + 1 To dictionary_object.Count - 1
      '比較
      If varData(i) > varData(j) Then
        varTemp = varData(i)
        varData(i) = varData(j)
        varData(j) = varTemp
      End If
    Next
  Next
  
  KeySort = varData
End Function

'dictionaryオブジェクトのキーをソートしてkey1=value1&key2=valu2...の文字列を返す
Private Function UrlParse(dictionary_object As Object) As String
  Dim strReqData As String
  Dim d As Variant
  Dim i As Long
  On Error Resume Next
  d = KeySort(dictionary_object)
  For i = 0 To UBound(d)
    strReqData = strReqData & "&" & CStr(d(i)) & "=" & dictionary_object(d(i))
  Next
  If Err.Number = 0 Then
    UrlParse = Mid(strReqData, 2)
  Else
    UrlParse = ""
  End If
  On Error GoTo 0
End Function

Private Function MakeSHA1Hash(bytValue() As Byte) As String
  Dim ret As Long
  Dim lngProv As Long 'コンテナオブジェクト
  Dim lngHash As Long 'ハッシュオブジェクト
  Dim lngHashSize As Long 'ハッシュサイズ
  Dim bytBuff() As Byte 'ハッシュが格納されるエリア
  Dim strHex As String
  Dim i As Long
  Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
  Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0"
  Const ALG_TYPE_ANY As Long = 0
  Const ALG_CLASS_HASH As Long = 32768
  Const ALG_TYPE_BLOCK As Long = 1536
  Const ALG_SID_SHA As Long = 4
  Const ALG_SID_SHA1 As Long = ALG_SID_SHA
  Const CALG_SHA As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA
  Const CALG_SHA1 As Long = CALG_SHA
  Const PROV_RSA_FULL As Long = 1
  Const HP_HASHVAL As Long = 2
  'SHA1ハッシュ
  Const lngHashType = CALG_SHA1
  
  'キーコンテナの作成
  ret = CryptAcquireContext(lngProv, vbNullString, MS_DEF_PROV, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
  If ret = False Then
    MakeSHA1Hash = ""
    Exit Function
  End If
  
  'ハッシュオブジェクトの作成
  ret = CryptCreateHash(lngProv, lngHashType, 0, 0, lngHash)
  If ret = False Then
    Call CryptReleaseContext(lngProv, 0)
    Exit Function
  End If
  
  'ハッシュデータを作る
  ret = CryptHashData(lngHash, bytValue(0), UBound(bytValue) + 1, 0)
  If ret = False Then
    GoTo ExitHandler
  End If
  
  '必要なサイズを取得
  ret = CryptGetHashParam(lngHash, HP_HASHVAL, ByVal 0, lngHashSize, 0)
  If ret = False Then
    GoTo ExitHandler
  End If
  
  'ハッシュを取り出す
  ReDim bytBuff(lngHashSize)
  For i = 0 To UBound(bytBuff)
    bytBuff(i) = 0
  Next
  ret = CryptGetHashParam(lngHash, HP_HASHVAL, bytBuff(0), lngHashSize, 0)
  If ret = False Then
    GoTo ExitHandler
  End If
  
  'HEX文字列へ
  For i = 0 To UBound(bytBuff) - 1
    strHex = strHex & Right("0" & LCase(Hex(bytBuff(i))), 2)
  Next
  
  MakeSHA1Hash = strHex
ExitHandler:
    Call CryptDestroyHash(lngHash)
    Call CryptReleaseContext(lngProv, 0)

End Function

'暗号化
Private Function hmac(ByVal key As String, ByVal data As String) As String
  Dim i As Integer
  Dim hash As String
  Dim key_byte() As Byte
  Dim key_len As Long
  Dim data_len As Long
  Dim ipad(63) As Byte
  Dim opad(63) As Byte
  Dim key_hash() As Byte
  Dim data_hash As String

  If key = "" And data = "" Then Exit Function

  key_len = Len(key)

  key_byte = StrConv(key, vbFromUnicode)
  If key_len > 64 Then
      key_hash = StrToBynary(MakeSHA1Hash(key_byte))
      key_len = 20
  Else
      key_hash = key_byte
  End If
  
  ReDim Preserve key_hash(63)
  For i = key_len To 63
    key_hash(i) = 0
  Next

  For i = 0 To 63
    ipad(i) = 0
    opad(i) = 0
  Next

  For i = 0 To 63
    ipad(i) = key_hash(i) Xor &H36
    opad(i) = key_hash(i) Xor &H5C
  Next

  data_hash = MakeSHA1Hash(CStr(ipad) & StrConv(data, vbFromUnicode))

  hash = MakeSHA1Hash(CStr(opad) & CStr(StrToBynary(data_hash)))

  hmac = hash
End Function

'バイト文字列からバイト配列を返す
Private Function StrToBynary(strHexString As String) As Byte()
  Dim buf() As Byte
  Dim i As Long
  
  ReDim Preserve buf(Len(CStr(strHexString)) \ 2 - 1)
  For i = 0 To Len(CStr(strHexString)) \ 2 - 1
    buf(i) = CByte("&h" & Mid(CStr(strHexString), i * 2 + 1, 2))
  Next
  StrToBynary = buf
End Function

'TwitterAPIの作成日から日付型の変数を返す
Private Function ConvertCreateTime(strCreated_at As String) As Date
  ConvertCreateTime = DateValue(Mid(strCreated_at, 5, 6) & Right(strCreated_at, 5)) + TimeValue(Mid(strCreated_at, 11, 9)) + TimeValue("09:00")
End Function

'TwitterAPIのレスポンスからTokenを抜き出す
Private Function GetOAuthToken(strTarget As String) As String
  Dim s, a, v
  s = Split(strTarget, "&")
  For Each a In s
    v = Split(a, "=")
    If v(0) = "oauth_token" Then
      GetOAuthToken = v(1)
      Exit Function
    End If
  Next
  GetOAuthToken = ""
End Function

'TwitterAPIのレスポンスからsecretを抜き出す
Private Function GetOAuthToken_secret(strTarget As String) As String
  Dim s, a, v
  s = Split(strTarget, "&")
  For Each a In s
    v = Split(a, "=")
    If v(0) = "oauth_token_secret" Then
      GetOAuthToken_secret = v(1)
      Exit Function
    End If
  Next
  GetOAuthToken_secret = ""
End Function

'HTMLエスケープ
Private Function html_escape(strString As String) As String
  Dim strTemp As String
  strTemp = strString
  strTemp = Replace(strTemp, "&amp;", "&")
  strTemp = Replace(strTemp, "&lt;", "<")
  strTemp = Replace(strTemp, "&gt;", ">")
  strTemp = Replace(strTemp, "&quot;", """")
  strTemp = Replace(strTemp, vbLf, "") 'ついでに改行も削除
  html_escape = strTemp
End Function

'アクセストークンをファイル名をフルパスで返す
Private Function GetTokenFileName() As String
  Dim strFileName As String
  Dim i As Long
  
  strFileName = ""
  i = InStr(1, ThisWorkbook.Name, ".")
  If i Then
    strFileName = Environ("USERPROFILE") & "\." & Mid(ThisWorkbook.Name, 1, InStr(1, ThisWorkbook.Name, ".") - 1)
  Else
    strFileName = Environ("USERPROFILE") & "\." & ThisWorkbook.Name
  End If
  GetTokenFileName = strFileName
End Function

'アクセストークンを読み出す。
'戻り値はバリアント配列(access_token,access_token_secret)。エラーならゼロを返す
Private Function GetToken() As Variant
  Dim intFileNo As Integer
  Dim strFileName As String
  Dim strData(1) As String
  
  On Error GoTo ErrorHandler
  strFileName = GetTokenFileName()
  If strFileName = "" Then
    GetToken = 0
    Exit Function
  End If
  If Dir(strFileName) = "" Then
    GetToken = 0
    Exit Function
  End If
  
  intFileNo = FreeFile()
  Open strFileName For Input As #intFileNo
  Input #intFileNo, strData(0)
  Input #intFileNo, strData(1)
  Close #intFileNo
  GetToken = strData
  Exit Function
ErrorHandler:
  GetToken = 0
End Function

'アクセストークンを保存する。エラーならFalseを返す
Private Function SaveToken(access_token As String, access_token_secret As String) As Boolean
  Dim intFileNo As Integer
  Dim strFileName As String
  
  On Error GoTo ErrorHandler
  strFileName = GetTokenFileName()
  If strFileName = "" Then
    SaveToken = False
    Exit Function
  End If
  
  intFileNo = FreeFile()
  Open strFileName For Output As #intFileNo
  Print #intFileNo, access_token
  Print #intFileNo, access_token_secret
  Close #intFileNo
  SaveToken = True
  Exit Function
ErrorHandler:
  SaveToken = False
End Function
