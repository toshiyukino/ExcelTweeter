Attribute VB_Name = "ExcelTweeter"
Option Explicit

'-----------------------------------
'定数
'-----------------------------------
Public Enum TimeLineName
  home_timeline = 1
  friends_timeline = 2
  user_timeline = 3
  replies = 4
  mentions = 5
  retweeted_by_me = 6
  retweeted_to_me = 7
  retweets_of_me = 8
  favorites_timeline = 20
End Enum
Public Enum TweetType
  Default_Tweet = 1 '普通のポスト
  Reply_Tweet = 2   '返信
  Re_Tweet = 3      '公式リツート
  Rt_Tweet = 4      '非公式リツイート
  Qt_Tweet = 5      '引用ツイート
  Dm_Post = 6       'ダイレクトメッセージ送信
  Fv_Post = 7       'お気に入り登録
End Enum
Public Enum DelType
  status_delete = 1 'ツイート削除
  dm_delete = 2     'ダイレクトメッセージ削除
  fv_delete = 3     'ダイレクトメッセージ削除
End Enum

'-----------------------------------
'ユーザ定義型
'-----------------------------------
Public Type StatusUser
  Id As String
  Name As String
  ScreenName As String
  ProfileImageUrl As String
End Type

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
Private Const retw_url = "https://api.twitter.com/1/statuses/retweet/" ' & "statusid.xml"
Private Const timeline_url = "https://api.twitter.com/1/statuses/"
Private Const twtshow_url = "https://api.twitter.com/1/statuses/show/" ' & "statusid.xml"
Private Const fav_url = "https://api.twitter.com/1/favorites.xml"
Private Const fav_add = "https://api.twitter.com/1/favorites/create/"  ' & "id.format"
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
  (ByVal hProv As Long, ByVal algid As Long, ByVal hKey As Long, _
  ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" _
  (ByVal hHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" _
  (ByVal hHash As Long, ByRef pbData As Any, _
  ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32.dll" _
  (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Any, _
  ByRef pdwDataLen As Long, ByVal dwFlags As Integer) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" _
  (ByVal hProv As Long, ByVal algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" _
  (ByVal hKey As Long) As Long
Private Declare Function CryptSetHashParam Lib "advapi32.dll" _
  (ByVal hHash As Long, ByVal dwParam As Long, _
  ByRef pbData As Any, ByVal dwFlags As Integer) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" _
  (ByVal hProv As Long, ByRef pbData As Any, _
  ByVal dwDataLen As Long, ByVal hPubKey As Long, _
  ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" _
  (Destination As Any, ByVal Length As Long)
Private Declare Function CryptBinaryToString Lib "crypt32.dll" Alias "CryptBinaryToStringA" _
  (ByRef pbBinary As Any, ByVal cbBinary As Long, _
  ByVal dwFlags As Long, ByVal pszString As String, _
  ByRef pcchString As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
  (ByVal hwnd As Long, ByVal lpOperation As String, _
  ByVal lpFile As String, ByVal lpParameters As String, _
  ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" _
  (ByVal dwMilliseconds As Long)
Private Type HMAC_Info
  HashAlgid As Long
  pbInnerString As Byte
  cbInnerString As Long
  pbOuterString As Byte
  cbOuterString As Long
End Type
Private Type BLOBHEADER
  bType As Byte
  bVersion As Byte
  reserved As Integer
  aiKeyAlg As Long
End Type
Private Type key_blob
  hdr As BLOBHEADER
  len As Long
  key(1024) As Byte '// TODO might want to dynamically allocate this, Should Be Fine though
End Type


'-----------------------------------------------------------------------
' Public method
'-----------------------------------------------------------------------

'アクセストークンを保存しているファイルを削除
'成功すると True を返す
Public Function DelTokenFile() As Boolean
  Dim intFileNo As Integer
  Dim strFileName As String
  
  On Error GoTo ErrorHandler
  
  ' 設定ファイルのパスを取得
  strFileName = GetTokenFileName()
  
  ' 取得できなかったらエラー
  If strFileName = "" Then
    DelTokenFile = False
    Exit Function
  End If
  
  ' 削除
  Kill strFileName
  DelTokenFile = True
  Exit Function

ErrorHandler:
  DelTokenFile = False
End Function

'成功すると「ok」を返す
Public Function TweetPost( _
    strPost As String, _
    Optional Tweet_Type As TweetType = Default_Tweet, _
    Optional strStatusID As String = "" _
) As String
  
  Dim xhr As Object 'MSXML2.ServerXMLHTTP60
  Dim param As Object  'Scripting.Dictionary
  Dim strReqURL As String
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
  
  'ポストの種類で処理分岐
  Select Case Tweet_Type
    
    '公式リツイート
    Case Re_Tweet
      strReqURL = retw_url & strStatusID & ".xml"
      If strStatusID = "" Then
        MsgBox "リツート元のステータスＩＤが取得できませんでした。", vbCritical
        TweetPost = "error"
        Exit Function
      End If
      param("id") = strStatusID
            
    'お気に入りに登録
    Case Fv_Post
      strReqURL = fav_add & strStatusID & ".xml"
      If strStatusID = "" Then
        MsgBox "リツート元のステータスＩＤが取得できませんでした。", vbCritical
        TweetPost = "error"
        Exit Function
      End If
      param("id") = strStatusID
    
    'その他
    Case Else
      strReqURL = post_url
      param("status") = UrlEncode(Left(strPost, 140)) '140文字
      '通常のポスト以外はは返信元ＩＤを入れる
      If Tweet_Type <> Default_Tweet And strStatusID <> "" Then
        param("in_reply_to_status_id") = strStatusID
      End If
  
  End Select
  
  'シグネチャ作成
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
  Do Until xhr.readyState = 4
    DoEvents
  '  Sleep 100
  Loop
  
  'ステータスがエラーの場合
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

'成功すると「ok」を返す
Public Function StatusShow( _
  strStatusID As String _
) As String
  Dim xhr As Object 'MSXML2.ServerXMLHTTP60
  Dim param As Object  'Scripting.Dictionary
  Dim strReqURL As String
  Dim XMLDOM As Object 'MSXML2.DOMDocument
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim i As Long

  If Not IsArray(GetToken) Then
    If isOAuth = False Then
      MsgBox "アクセストークンが取得できませんでした", vbCritical
      StatusShow = "error"
      Exit Function
    End If
  End If
  
  atoken = GetToken(0)
  atoken_secret = GetToken(1)
  
  Set param = CreateObject("Scripting.Dictionary")
  param("oauth_token") = atoken
  strReqURL = twtshow_url & strStatusID & ".xml"
  strSig = MakeSignature("GET", strReqURL, param, UrlEncode(Consumer_secret) & "&" & UrlEncode(atoken_secret))
  param("oauth_signature") = UrlEncode(strSig)
  Set xhr = CreateRequest("GET", strReqURL & "?" & UrlParse(param))
  If xhr Is Nothing Then
    MsgBox "リクエストオブジェクトが作成できませんでした", vbCritical
    StatusShow = "error"
    Exit Function
  End If
  xhr.send
  
  '読み込みが完了するまでループ
  Do Until xhr.readyState = 4
    DoEvents
  '  Sleep 100
  Loop
  
  If xhr.Status <> 200 Then
    MsgBox xhr.getAllResponseHeaders
    Set XMLDOM = xhr.responseXML
    If XMLDOM Is Nothing Then
      'MsgBox xhr.ResponseText
      StatusShow = "error"
    Else
      StatusShow = XMLDOM.selectSingleNode("hash/error").FirstChild.nodeValue
    End If
    Exit Function
  Else
    StatusShow = "ok"
    Debug.Print xhr.responseText
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
    Case favorites_timeline
      strTL_url = fav_url
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
  Do Until xhr.readyState = 4
    DoEvents
    'Sleep 100
  Loop
 
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
'Debug.Print objStatus.selectSingleNode("user/profile_image_url").FirstChild.nodeValue
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


'-----------------------------------------------------------------------
' Private method
'-----------------------------------------------------------------------

'コンシューマキーを使ってアクセストークンを取得してファイルに保存する。
Private Function isOAuth() As Boolean
  Dim xhr As Object 'MSXML2.ServerXMLHTTP60
  Dim dicParam As Scripting.Dictionary
  Dim strRes As String
  Dim otoken As String, otoken_secret As String
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim strPin As String
  Dim i As Long
  
  On Error GoTo ErrorHandler
  
  Set dicParam = CreateObject("Scripting.Dictionary")

  strSig = MakeSignature("GET", reqt_url, dicParam, UrlEncode(Consumer_secret) & "&")
  param("oauth_signature") = UrlEncode(strSig)

  Set xhr = CreateRequest("GET", reqt_url & "?" & UrlParse(dicParam))
  If xhr Is Nothing Then
    MsgBox "リクエストオブジェクトが作成できませんでした", vbCritical
    isOAuth = False
    Exit Function
  End If
  xhr.send '送信
  
  '取得失敗
  If xhr.Status <> 200 Then
    isOAuth = False
    Exit Function
  End If
  
  'レスポンステキスト
  strRes = xhr.responseText
  
  'authトークン（一時的に使う）
  otoken = GetOAuthToken(strRes)
  otoken_secret = GetOAuthToken_secret(strRes)

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

  dicParam.RemoveAll
  dicParam("oauth_verifier") = strPin '今回だけ（PINコード）
  dicParam("oauth_token") = otoken '今回だけ（authトークン）
  strSig = MakeSignature("GET", acct_url, dicParam, UrlEncode(Consumer_secret) & "&")
  dicParam("oauth_signature") = UrlEncode(strSig)

  Set xhr = CreateRequest("GET", acct_url & "?" & UrlParse(dicParam))
  xhr.send
  
  If xhr.Status <> 200 Then
    isOAuth = False
    Exit Function
  End If

  'レスポンステキスト
  strRes = xhr.responseText
  
  atoken = GetOAuthToken(strRes)
  atoken_secret = GetOAuthToken_secret(strRes)

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
  
ErrorHandler:
  isOAuth = False 'エラー
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
  'Call xhr.Open(strMethod, strRequestParm, True)  '非同期処理
  
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
  'エンコードされないので文字の対策
  s = Replace(s, "(", "%28")  '(
  s = Replace(s, ")", "%29")  ')
  s = Replace(s, "!", "%21")  '!
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

'暗号化
Private Function hmac(ByVal strKey As String, ByVal strData As String) As String
  Dim bytKey() As Byte
  Dim bytData() As Byte
  Dim ret As Long
  Dim lngProv As Long       'コンテナオブジェクト
  Dim lngHash As Long       'ハッシュオブジェクト
  Dim lngHmacHash As Long   'ハッシュオブジェクト
  Dim lngHashSize As Long   'ハッシュサイズ
  Dim lngKey As Long        'キーオブジェクト
  Dim bytBuff() As Byte     'ハッシュが格納されるエリア
  Dim strHex As String      '16進数文字列
  Dim i As Long
  Dim HmacInfo As HMAC_Info
  Dim keyblob As key_blob
  Dim key_len As Long
  Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
  Const MS_DEF_PROV As String = "Microsoft Base Cryptographic Provider v1.0"
  Const ALG_TYPE_ANY As Long = 0
  Const ALG_CLASS_HASH As Long = 32768
  Const ALG_TYPE_BLOCK As Long = 1536
  Const ALG_SID_SHA As Long = 4
  Const ALG_SID_SHA1 As Long = ALG_SID_SHA
  Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
  Const ALG_TYPE_STREAM As Long = 2048
  Const ALG_SID_RC4 As Long = 1
  Const ALG_SID_RC2 As Long = 2
  Const ALG_SID_HMAC As Long = 9
  Const CALG_SHA As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA
  Const CALG_SHA1 As Long = CALG_SHA
  Const CALG_RC2 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2
  Const CALG_RC4 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4
  Const CALG_HMAC As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_HMAC
  Const HP_HMAC_INFO = &H5
  Const HP_HASHVAL As Long = 2
  Const PROV_RSA_FULL As Long = 1
  Const PLAINTEXTKEYBLOB As Long = 8
  Const CUR_BLOB_VERSION As Long = 2
  Const CRYPT_IPSEC_HMAC_KEY = &H100
  
  If strKey = "" And strData = "" Then Exit Function
  
  hmac = ""
  strHex = ""
  
  'バイト配列へ
  bytKey = StrConv(strKey, vbFromUnicode)
  bytData = StrConv(strData, vbFromUnicode)

  '1024バイトチェック
  key_len = UBound(bytKey) + 1
  If key_len > 1024 Then
    hmac = ""
    Exit Function
  End If

  'キーコンテナの作成
  ret = CryptAcquireContext(lngProv, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
  If ret = False Then
    GoTo ExitHandler
  End If

'  '鍵作り
'  ret = CryptDeriveKey(lngProv, CALG_RC2, lngHash, 0, lngKey)
'  If ret = False Then
'    Call CryptDestroyKey(lngKey)
'    GoTo ExitHandler
'  End If

  '// key creation based on
  '// http://mirror.leaseweb.com/NetBSD/NetBSD-release-5-0/src/dist/wpa/src/crypto/crypto_cryptoapi.c
  keyblob.hdr.bType = PLAINTEXTKEYBLOB
  keyblob.hdr.bVersion = CUR_BLOB_VERSION
  keyblob.hdr.reserved = 0
  '/*
  '* Note: RC2 is not really used, but that can be used to
  '* import HMAC keys of up to 16 byte long.
  '* CRYPT_IPSEC_HMAC_KEY flag for CryptImportKey() is needed to
  '* be able to import longer keys (HMAC-SHA1 uses 20-byte key).
  '*/
  keyblob.hdr.aiKeyAlg = CALG_RC2
  keyblob.len = key_len
  Call ZeroMemory(keyblob.key(0), key_len)
  Call CopyMemory(keyblob.key(0), bytKey(0), key_len)
  ret = CryptImportKey(lngProv, keyblob, 12 + key_len, 0, CRYPT_IPSEC_HMAC_KEY, lngKey)
  If ret = False Then
    GoTo ExitHandler
  End If
  
  'ハッシュオブジェクトの作成
  ret = CryptCreateHash(lngProv, CALG_HMAC, lngKey, 0, lngHmacHash)
  If ret = False Then
    GoTo ExitHandler
  End If
  
  'パラメータセット
  HmacInfo.HashAlgid = CALG_SHA1
  ret = CryptSetHashParam(lngHmacHash, HP_HMAC_INFO, HmacInfo, 0)
  If ret = False Then
    GoTo ExitHandler
  End If

  'ハッシュデータを作る
  ret = CryptHashData(lngHmacHash, bytData(0), UBound(bytData) + 1, 0)
  If ret = False Then
    GoTo ExitHandler
  End If

  '必要なサイズを取得
  ret = CryptGetHashParam(lngHmacHash, HP_HASHVAL, ByVal 0, lngHashSize, 0)
  If ret = False Then
    GoTo ExitHandler
  End If
  
  'ハッシュを取り出す
  ReDim bytBuff(lngHashSize - 1)
  For i = 0 To UBound(bytBuff)
    bytBuff(i) = 0
  Next
  ret = CryptGetHashParam(lngHmacHash, HP_HASHVAL, bytBuff(0), lngHashSize, 0)
  If ret = False Then
    GoTo ExitHandler
  End If

  'HEX文字列へ
  For i = 0 To UBound(bytBuff)
    strHex = strHex & Right("0" & LCase(Hex(bytBuff(i))), 2)
  Next
  
ExitHandler:
  If (lngHmacHash) Then
    CryptDestroyHash (lngHmacHash)
  End If
  If (lngKey) Then
    Call CryptDestroyKey(lngKey)
  End If
  If (lngHash) Then
    Call CryptDestroyHash(lngHash)
  End If
  If (lngProv) Then
    Call CryptReleaseContext(lngProv, 0)
  End If
  hmac = strHex
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
  If Not IsArray(s) Then: GoTo ErrorHandler
  For Each a In s
    v = Split(a, "=")
    If Not IsArray(v) Then: GoTo ErrorHandler
    If v(0) = "oauth_token" Then
      GetOAuthToken = v(1)
      Exit Function
    End If
  Next
ErrorHandler:
  GetOAuthToken = ""
End Function

'TwitterAPIのレスポンスからsecretを抜き出す
Private Function GetOAuthToken_secret(strTarget As String) As String
  Dim s, a, v
  s = Split(strTarget, "&")
  If Not IsArray(s) Then: GoTo ErrorHandler
  For Each a In s
    v = Split(a, "=")
    If Not IsArray(v) Then: GoTo ErrorHandler
    If v(0) = "oauth_token_secret" Then
      GetOAuthToken_secret = v(1)
      Exit Function
    End If
  Next
ErrorHandler:
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
  strTemp = Replace(strTemp, vbLf, " ") 'ついでに改行も削除
  html_escape = strTemp
End Function

'アクセストークンをファイル名をフルパスで返す
Private Function GetTokenFileName() As String
  Dim strFileName As String
  Dim i As Long
  Dim strHomePath As String
  
  On Error GoTo ErrorHandler
  
  strFileName = ""
  
  If Len(Environ("USERPROFILE")) = 0 Then
    strHomePath = CurDir
  Else
    strHomePath = Environ("USERPROFILE")
  End If
  
  i = InStr(1, ThisWorkbook.Name, ".")
  If i Then
    strFileName = strHomePath & "\." & Mid(ThisWorkbook.Name, 1, InStr(1, ThisWorkbook.Name, ".") - 1)
  Else
    strFileName = strHomePath & "\." & ThisWorkbook.Name
  End If
  
  GetTokenFileName = strFileName
  Exit Function

ErrorHandler:
  GetTokenFileName = ""
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
