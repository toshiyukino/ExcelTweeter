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
  fv_delete = 3     'お気に入り削除
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
Private Const UserAgent = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1;" & _
                          ".NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.1) "
'-----------------------------------
'Temporary Internet Files
'-----------------------------------
Private Const UseImageFile = True 'イメージを使うか
Private Const IeTempInternetFiles = "" '(空白でIEと同じ場所)
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
    Optional strStatusId As String = "" _
) As String
  
  Dim res As Object
  Dim param As Object  'Scripting.Dictionary
  Dim strReqURL As String
  Dim XMLDOM As Object 'MSXML2.DOMDocument
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim i As Long

  If GetToken Is Nothing Then
    If isOAuth = False Then
      MsgBox "アクセストークンが取得できませんでした", vbCritical
      TweetPost = ""
      Exit Function
    End If
  End If
  
  atoken = GetToken(1)
  atoken_secret = GetToken(2)
  
  Set param = CreateObject("Scripting.Dictionary")
  param("oauth_token") = atoken
  param("source") = "ExcelTweeter"
  
  'ポストの種類で処理分岐
  Select Case Tweet_Type
    
    '公式リツイート
    Case Re_Tweet
      strReqURL = retw_url & strStatusId & ".xml"
      If strStatusId = "" Then
        MsgBox "リツート元のステータスＩＤが取得できませんでした。", vbCritical
        TweetPost = "error"
        Exit Function
      End If
      param("id") = strStatusId
            
    'お気に入りに登録
    Case Fv_Post
      strReqURL = fav_add & strStatusId & ".xml"
      If strStatusId = "" Then
        MsgBox "リツート元のステータスＩＤが取得できませんでした。", vbCritical
        TweetPost = "error"
        Exit Function
      End If
      param("id") = strStatusId
    
    'その他
    Case Else
      strReqURL = post_url
      param("status") = strPost '140文字
      '通常のポスト以外はは返信元ＩＤを入れる
      If Tweet_Type <> Default_Tweet And strStatusId <> "" Then
        param("in_reply_to_status_id") = strStatusId
      End If
  
  End Select
  
  '署名を作成
  strSig = MakeSignature("POST", strReqURL, param, UrlEncode(Consumer_secret) & "&" & UrlEncode(atoken_secret))
  param("oauth_signature") = strSig
    
  Set res = HttpOpen("POST", strReqURL, UrlParse(param))
  If res Is Nothing Then
    MsgBox "リクエストに失敗しました", vbCritical
    TweetPost = 0
    Exit Function
  End If
  
  'ステータスがエラーの場合
  If res("Status") <> 200 Then
    MsgBox res("getAllResponseHeaders")
    Set XMLDOM = res("responseXML")
    If XMLDOM Is Nothing Then
      TweetPost = "error"
    Else
      TweetPost = XMLDOM.selectSingleNode("hash/error").FirstChild.nodeValue
    End If
    Exit Function
  Else
    TweetPost = "ok"
  End If
  
  Set res = Nothing
  Set param = Nothing
  Set XMLDOM = Nothing
End Function

'成功すると「ok」を返す
Public Function StatusShow( _
  strStatusId As String _
) As String
  Dim res As Object     'Collection
  Dim param As Object   'Scripting.Dictionary
  Dim strReqURL As String
  Dim XMLDOM As Object  'MSXML2.DOMDocument
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim i As Long

  If GetToken Is Nothing Then
    If isOAuth = False Then
      'MsgBox "アクセストークンが取得できませんでした", vbCritical
      StatusShow = "Error / アクセストークンが取得できませんでした"
      Exit Function
    End If
  End If
  
  atoken = GetToken(1)
  atoken_secret = GetToken(2)
  
  Set param = CreateObject("Scripting.Dictionary")
  param("oauth_token") = atoken
  strReqURL = twtshow_url & strStatusId & ".xml"
  strSig = MakeSignature("GET", strReqURL, param, UrlEncode(Consumer_secret) & "&" & UrlEncode(atoken_secret))
  param("oauth_signature") = strSig
  
  Set res = HttpOpen("GET", strReqURL & "?" & UrlParse(param))
  If res Is Nothing Then
    MsgBox "リクエストオブジェクトが作成できませんでした", vbCritical
    StatusShow = "error"
    Exit Function
  End If
  
  If res("Status") <> 200 Then
    MsgBox xhr.getAllResponseHeaders
    Set XMLDOM = xhr.responseXML
    If XMLDOM Is Nothing Then
      Debug.Print res("Status") & ":" & res("StatusText")
      'MsgBox xhr.ResponseText
      StatusShow = "error"
    Else
      StatusShow = XMLDOM.selectSingleNode("hash/error").FirstChild.nodeValue
    End If
    Exit Function
  Else
    StatusShow = "ok"
  End If
End Function


'(x,0):id
'(x,1):screen name
'(x,2):tweet text
'(x,3):create time
'(x,4):image file
'引数 timeline_count:取得するタイムラインの数
'     timeline_name:タイムラインの種類
'     otpdic:その他指定項目をハッシュで渡す
'戻り値が配列か確認して使う
Public Function GetTimeLine _
  (Optional timeline_count As Long = 20, _
  Optional timeline_name As TimeLineName = home_timeline, _
  Optional optdic As Object = Nothing _
) As Variant
  Dim res As Object
  Dim param As Object  'Scripting.Dictionary
  Dim XMLDOM As Object 'MSXML2.DOMDocument
  Dim Statuses As Object 'MSXML2.IXMLDOMNode
  Dim objStatus As Object 'MSXML2.IXMLDOMElement
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim i As Long, j As Long
  Dim strTimeLine() As String
  Dim strTemp As String
  Dim strTL_url As String
  Dim dickey As Variant       'scripting.dictionary key
  Dim urls As Object 'entities
  Dim res_img As Object
  Dim bytimg() As Byte
  Dim strImgPath As String
  
  On Error GoTo ErrorHandler
  
  If GetToken Is Nothing Then
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
  
  atoken = GetToken(1)
  atoken_secret = GetToken(2)

  Set param = CreateObject("Scripting.Dictionary")
  param("oauth_token") = atoken 'Access_Token
  param("count") = CStr(timeline_count)
  param("include_entities") = "true"  '追加
  If Not optdic Is Nothing Then  'オプションがあれば追加
    For Each dickey In optdic.Keys
      param(dickey) = optdic(dickey)
    Next
  End If
  strSig = MakeSignature("GET", strTL_url, param, UrlEncode(Consumer_secret) & "&" & UrlEncode(atoken_secret))
  param("oauth_signature") = strSig
  
  Set res = HttpOpen("GET", strTL_url & "?" & UrlParse(param))
  If res Is Nothing Then
    MsgBox "リクエストに失敗しました", vbCritical
    GoTo ErrorHandler
  End If
  
  If res("Status") <> 200 Then
    Set XMLDOM = res("responseXML")
    If Not XMLDOM Is Nothing Then
      Debug.Print XMLDOM.selectSingleNode("hash/error").FirstChild.nodeValue
      GetTimeLine = 0
      Exit Function
    Else
      GoTo ErrorHandler
    End If
  End If
  
  ReDim strTimeLine(param("count") - 1, 4)
  'Set XMLDOM = res("responseXML")
  Set XMLDOM = CreateObject("MSXML2.DOMDocument")
  If XMLDOM.loadXML(Replace(Replace(res("responseText"), vbLf, ""), vbCr, "")) = False Then
    GoTo ErrorHandler
  End If
  
  XMLDOM.setProperty "SelectionLanguage", "XPath"
  Set Statuses = XMLDOM.childNodes(1)
  If Statuses Is Nothing Then
    GoTo ErrorHandler
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
    
    'プロフィールイメージ(リファクタリングの必要あり)
    If UseImageFile = True Then
      strTimeLine(i, 4) = objStatus.selectSingleNode("user/profile_image_url_https").FirstChild.nodeValue
      strImgPath = GetIeTempInternetFiles & strTimeLine(i, 1) & Mid(strTimeLine(i, 4), InStrRev(strTimeLine(i, 4), "."), 4)
      
      '既に保存してあるものはダウンロードしない
      If Dir(strImgPath) = "" Then
        'プロフィールイメージをダウンロード
        Set res_img = HttpOpen("GET", strTimeLine(i, 4))
        If Not res_img Is Nothing Then
          '書き込み
          Open strImgPath For Binary As #1
            bytimg() = res_img("responseBody")
            Put #1, , bytimg()
          Close #1
          Set res_img = Nothing
          Debug.Print "download..", strImgPath
        End If
      End If
      strTimeLine(i, 4) = strImgPath
    End If
    
    't.co対策
    If objStatus.selectNodes("entities/urls/url").Length > 0 Then  'entities/urlsに入るのであるか確認
      'urlsを取得（複数ある）
      Set urls = objStatus.selectNodes("entities/urls")
      For j = 0 To urls.Length - 1  'すべて取得する
        strTimeLine(i, 2) = Replace(strTimeLine(i, 2), urls(j).selectSingleNode("url/url").FirstChild.nodeValue, urls(j).selectSingleNode("url/expanded_url").FirstChild.nodeValue)
      Next
    End If
    
    
    i = i + 1
  Next
  GetTimeLine = strTimeLine
  Set res = Nothing
  Set XMLDOM = Nothing
  
  Exit Function
ErrorHandler:
  GetTimeLine = 0
  If Not res Is Nothing Then
    Debug.Print res("responseText")
  End If
  Debug.Print Err.Number; Err.Description
  Close #1
End Function


'-----------------------------------------------------------------------
' Private method
'-----------------------------------------------------------------------

'コンシューマキーを使ってアクセストークンを取得してファイルに保存する。
Private Function isOAuth() As Boolean
  Dim res As Object
  Dim dicParam As Object 'Scripting.Dictionary
  Dim strRes As String
  Dim otoken As String, otoken_secret As String
  Dim atoken As String, atoken_secret As String
  Dim strSig As String
  Dim strPin As String
  Dim i As Long
  
  On Error GoTo ErrorHandler
  
  Set dicParam = CreateObject("Scripting.Dictionary")

  strSig = MakeSignature("GET", reqt_url, dicParam, UrlEncode(Consumer_secret) & "&")
  dicParam("oauth_signature") = strSig

  Set res = HttpOpen("GET", reqt_url & "?" & UrlParse(dicParam))
  If res Is Nothing Then
    isOAuth = False
    Exit Function
  End If
  
  '取得失敗
  If res("Status") <> 200 Then
    Debug.Print "authトークン取得エラー"
    Debug.Print res("Status") & ":" & res("StatusText")
    GoTo ErrorHandler
  End If
  
  'レスポンステキスト
  strRes = res("responseText")
  
  'authトークン（一時的に使う）
  otoken = GetOAuthToken(strRes)
  otoken_secret = GetOAuthToken_secret(strRes)

  'レスポンスにトークンが含まれていない場合
  If otoken = "" Or otoken_secret = "" Then
    Debug.Print "authトークン取得エラー"
    GoTo ErrorHandler
  End If

  'PIN取得の為ブラウザを起動（引数にauthトークンを指定）
  With CreateObject("WScript.Shell")
    .Run auth_url & "?oauth_token=" & otoken
  End With
  strPin = InputBox("pinを入力")
  If strPin = "" Then
    Debug.Print "pin入力エラー"
    GoTo ErrorHandler
  End If

  Set res = Nothing
  dicParam.RemoveAll
  dicParam("oauth_verifier") = strPin '今回だけ（PINコード）
  dicParam("oauth_token") = otoken    '今回だけ（authトークン）
  strSig = MakeSignature("GET", acct_url, dicParam, UrlEncode(Consumer_secret) & "&")
  dicParam("oauth_signature") = strSig

  Set res = HttpOpen("GET", acct_url & "?" & UrlParse(dicParam))
  
  If res("Status") <> 200 Then
    Debug.Print "アクセストークン取得エラー"
    Debug.Print res("Status") & ":" & res("StatusText")
    GoTo ErrorHandler
  End If

  'レスポンステキスト
  strRes = res("responseText")
  
  'アクセストークンの取得
  atoken = GetOAuthToken(strRes)
  atoken_secret = GetOAuthToken_secret(strRes)

  'レスポンスにトークンが含まれていない場合
  If atoken = "" Or atoken_secret = "" Then
    Debug.Print "アクセストークン取得エラー"
    Debug.Print res("Status") & ":" & res("StatusText")
    GoTo ErrorHandler
    Exit Function
  End If

  'ファイルに保存
  If SaveToken(atoken, atoken_secret) Then
    isOAuth = True  '保存成功
  Else
    isOAuth = False '保存エラー
  End If
  
  Exit Function
ErrorHandler:
  isOAuth = False 'エラー
End Function


'リクエストする（エラーは Nothing を返す）
Private Function HttpOpen(strMethod As String, strUrl As String, Optional strParam = "") As Object 'Collection
  Dim xhr As Object 'MSXML2.ServerXMLHTTP60
  Dim ua As String
  Dim col As Collection
  
  On Error GoTo ErrorHandler
  
  Set xhr = CreateObject("Msxml2.ServerXMLHTTP.6.0") 'MSXML2.ServerXMLHTTP60
  Set col = New Collection

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
  Call xhr.Open(strMethod, strUrl, False) '同期処理
  
  'プロキシ認証
  If Len(proxy_user) > 0 Then
    If InStr(1, proxy_user, ":") > 0 Then
      xhr.setProxyCredentials Split(proxy_user, ":")(0), Split(proxy_user, ":")(1)
    End If
  End If
  
  'リクエストヘッダーのセット
  If LCase(strMethod) = "post" Then
    Call xhr.setRequestHeader("Content-type", "application/x-www-form-urlencoded")
  End If
  Call xhr.setRequestHeader("User-Agent", ua)
  Call xhr.setRequestHeader("Pragma", "no-cache")
  Call xhr.setRequestHeader("Cache-Control", "Private")
  Call xhr.setRequestHeader("Expires", "-1")
  
  '送信
  If LCase(strMethod) = "get" Then
    xhr.send
  Else
    xhr.send strParam
  End If
  
  On Error Resume Next
  'Collectionへ入れる
  col.Add "", "RequestUrl"   'リクエストURL
  col.Add "", "Data"      'POSTリクエスト時のデータ
  col.Add Nothing, "responseBody"
  col.Add "", "responseText"
  col.Add Nothing, "responseXML"
  col.Add "", "getAllResponseHeaders"
  col.Add 0, "status"
  col.Add "", "statusText"
  
  '一回削除して入れ直す
  col.Remove "RequestUrl": col.Add strUrl, "RequestUrl"   'リクエストURL
  col.Remove "Data": col.Add strParam, "Data"       'POSTリクエスト時のデータ
  col.Remove "responseBody": col.Add xhr.responseBody, "responseBody"
  col.Remove "responseText": col.Add xhr.responseText, "responseText"
  col.Remove "responseXML": col.Add xhr.responseXML, "responseXML"
  col.Remove "getAllResponseHeaders": col.Add xhr.getAllResponseHeaders, "getAllResponseHeaders"
  col.Remove "status": col.Add xhr.Status, "status"
  col.Remove "statusText": col.Add xhr.statusText, "statusText"
  
  Set xhr = Nothing
  Set HttpOpen = col
  On Error GoTo 0
  Exit Function
ErrorHandler:
  Set HttpOpen = Nothing
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
  buf = hmac(strHmacKey, strReqData)
  MakeSignature = EncodeBase64(buf)
End Function

'UTF8バイト配列を返す
Private Function ConvUTF8(strValue)
  On Error Resume Next
  Dim UTF8_enc
  Dim bytBuff
  
  Set UTF8_enc = CreateObject("System.Text.UTF8Encoding")
  If Not UTF8_enc Is Nothing Then
    'バイト配列へ ('_4'の追加で、強制的に指定のオーバーロードメソッドを
    '呼び出す裏技(オーバーロードさせない))
    bytBuff = UTF8_enc.GetBytes_4(strValue)
  Else
    Set UTF8_enc = CreateObject("ADODB.Stream")
    With UTF8_enc
      .Open
      .Type = 2   'adTypeText
      .Charset = "utf-8"
      .WriteText = strValue
      .Position = 0
      .Type = 1   'adTypeBinary
      .Position = 3 'utf-8はオフセットさせる
      bytBuff = .Read()
      .Close
    End With
  End If
  ConvUTF8 = bytBuff
  Set UTF8_enc = Nothing
End Function

'英字, 数字, '-', '.', '_', '~' 以外をURL(パーセント)エンコードする
Private Function UrlEncode(ByVal strTarget)
  Dim s, b, tmp
  Dim i
  b = ConvUTF8(strTarget)
  For i = 0 To UBound(b)
    tmp = Chr(b(i))
    Select Case tmp
      Case "a" To "z", "A" To "Z", "0" To "9", "_", ".", "~", "-"
        s = s & tmp
      Case Else
        s = s & "%" & Right("0" & Hex(b(i)), 2)
    End Select
  Next
  UrlEncode = s
End Function

'JavaScriptの encodeURIComponent() クローン（但しスペースは'+'にする）
Private Function encodeURIComponent(ByVal strValue)
  Dim s, b, tmp
  Dim i
  b = ConvUTF8(strTarget)
  For i = 0 To UBound(b)
    tmp = Chr(b(i))
    Select Case tmp
      Case "a" To "z", "A" To "Z", "0" To "9", "'", "_", ".", "~", "/", "-", "*", "(", ")"
        s = s & tmp
      Case " "
        s = s & "+"
      Case Else
        s = s & "%" & Right("0" & Hex(b(i)), 2)
    End Select
  Next
  encodeURIComponent = s
End Function

'ActiveX版(Base64エンコード)
Private Function EncodeBase64(bytes)
  Dim dom, elm
  Set dom = CreateObject("Microsoft.XMLDOM")
  Set elm = dom.createElement("tmp")
  elm.DataType = "bin.base64"
  elm.nodeTypedValue = bytes
  EncodeBase64 = elm.Text
  Set dom = Nothing
  Set elm = Nothing
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
'key と value を encodeURIComponent でそれぞれエンコードする
Private Function UrlParse(dictionary_object As Object) As String
  Dim strReqData As String
  Dim d As Variant
  Dim i As Long
  On Error Resume Next
  d = KeySort(dictionary_object)
  For i = 0 To UBound(d)
    'strReqData = strReqData & "&" & encodeURIComponent(CStr(d(i))) & "=" & encodeURIComponent(dictionary_object(d(i)))
    'なぜかTwitterではエラーになるので対策
    strReqData = strReqData & "&" & UrlEncode(CStr(d(i))) & "=" & UrlEncode(dictionary_object(d(i)))
  Next
  If Err.Number = 0 Then
    UrlParse = Mid(strReqData, 2)
  Else
    UrlParse = ""
  End If
  On Error GoTo 0
End Function

'Hmac(SHA1)恐らくwin200以降なら動く
'バイナリを返すようにする(edit 2012/10/01)
Private Function hmac(ByVal strKey, ByVal strData)
  Dim HMACSHA1
  Dim bytUTF8
  Dim bytKey
  Dim bytData
  Dim bytBuff
  Dim i
  Dim strHex
  
  bytKey = ConvUTF8(strKey)
  bytData = ConvUTF8(strData)

  Set HMACSHA1 = CreateObject("System.Security.Cryptography.HMACSHA1")
  HMACSHA1.key = bytKey
  HMACSHA1.ComputeHash_2 (bytData)
  bytBuff = HMACSHA1.Hash
  
  Set HMACSHA1 = Nothing
  hmac = bytBuff
End Function

'TwitterAPIの作成日から日付型の変数を返す
Private Function ConvertCreateTime(strCreated_at As String) As Date
  ConvertCreateTime = DateValue(Mid(strCreated_at, 5, 6) & Right(strCreated_at, 5)) + TimeValue(Mid(strCreated_at, 11, 9)) + TimeValue("09:00")
End Function

'TwitterAPIのレスポンスからTokenを抜き出す
'エラーなら空白を返す
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
'エラーなら空白を返す
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

'アクセストークンを保存するファイル名をフルパスで返す
'エラーなら空白を返す
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
'戻り値はCollectionオブジェクト(access_token,access_token_secret)。
'エラーならFalseを返す
Private Function GetToken() As Variant
  Dim intFileNo As Integer
  Dim strFileName As String
  Dim strData(1) As String
  Dim col As Collection
  
  On Error GoTo ErrorHandler
  
  '設定ファイル名を取得
  strFileName = GetTokenFileName()
  If strFileName = "" Then
    GetToken = False
    Exit Function
  End If
  
  'ファイル存在確認
  If Dir(strFileName) = "" Then
    GetToken = False
    Exit Function
  End If
  
  'ファイル読み込み
  intFileNo = FreeFile()
  Open strFileName For Input As #intFileNo
    Input #intFileNo, strData(0)
    Input #intFileNo, strData(1)
  Close #intFileNo
  
  Set col = New Collection
  col.Add strData(0), "access_token"
  col.Add strData(1), "access_token_secret"
  Set GetToken = col
  
  Exit Function
ErrorHandler:
  GetToken = False
End Function

'アクセストークンを保存する。エラーならFalseを返す
Private Function SaveToken(access_token As String, access_token_secret As String) As Boolean
  Dim intFileNo As Integer
  Dim strFileName As String
  
  On Error GoTo ErrorHandler
  
  '設定ファイル名を取得
  strFileName = GetTokenFileName()
  If strFileName = "" Then
    SaveToken = False
    Exit Function
  End If
  
  '書き込み
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

'インターネットキャッシュフォルダを返す
Private Function GetIeTempInternetFiles()
  If IeTempInternetFiles = "" Then
    With CreateObject("Shell.Application")
      GetIeTempInternetFiles = Environ("userprofile") & "\Local Settings\" & .Namespace(32) & "\Content.IE5\ExcelTweeter\"
      If Dir(GetIeTempInternetFiles, vbDirectory) = "" Then
        With CreateObject("Scripting.FileSystemObject")
          Call .CreateFolder(GetIeTempInternetFiles)
          Exit Function
        End With
      End If
    End With
  Else
    GetIeTempInternetFiles = IeTempInternetFiles
  End If
End Function
