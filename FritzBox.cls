VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FritzBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' http://home.mengelke.de/cgi-bin/
'
' FritzBox Access via VBA
' by Andreas Schäfer 08/2018
' me@andreas-schaefer.com
'
' Adapted from http://home.mengelke.de/cgi-bin
'
' Copyright (c) 2017 Michael Engelke <http://www.MEngelke.de>, Andreas Schäfer 08/2018
' This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
' You should have received a copy of the GNU General Public License along with this program; if not, see <http://www.gnu.org/licenses/>.


Option Explicit

' Members
Private m_sid As String
Private m_host As String
Private m_page As String

Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32
Private m_lOnBits(30)
Private m_l2Power(30)


Private Sub Class_Initialize()
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)
    m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)
   
End Sub

Private Function LShift(lValue, iShiftBits)
 If iShiftBits = 0 Then
  LShift = lValue
  Exit Function
 ElseIf iShiftBits = 31 Then
  If lValue And 1 Then
   LShift = &H80000000
  Else
   LShift = 0
  End If
  Exit Function
 ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
  Err.Raise 6
 End If
 If (lValue And m_l2Power(31 - iShiftBits)) Then
  LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
 Else
  LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
 End If
End Function

Private Function RShift(lValue, iShiftBits)
 If iShiftBits = 0 Then
  RShift = lValue
  Exit Function
 ElseIf iShiftBits = 31 Then
  If lValue And &H80000000 Then
   RShift = 1
  Else
   RShift = 0
  End If
  Exit Function
 ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
  Err.Raise 6
 End If
 RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
 If (lValue And &H80000000) Then
  RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
 End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
 RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
 Dim lX4
 Dim lY4
 Dim lX8
 Dim lY8
 Dim lResult
 lX8 = lX And &H80000000
 lY8 = lY And &H80000000
 lX4 = lX And &H40000000
 lY4 = lY And &H40000000
 lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 If lX4 And lY4 Then
  lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
 ElseIf lX4 Or lY4 Then
  If lResult And &H40000000 Then
   lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
  Else
   lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
  End If
 Else
  lResult = lResult Xor lX8 Xor lY8
 End If
 AddUnsigned = lResult
End Function

Private Function f(x, y, z)
 f = (x And y) Or ((Not x) And z)
End Function

Private Function G(x, y, z)
 G = (x And z) Or (y And (Not z))
End Function

Private Function H(x, y, z)
 H = (x Xor y Xor z)
End Function

Private Function I(x, y, z)
 I = (y Xor (x Or (Not z)))
End Function

Private Sub FF(a, b, c, d, x, s, ac)
 a = AddUnsigned(a, AddUnsigned(AddUnsigned(f(b, c, d), x), ac))
 a = RotateLeft(a, s)
 a = AddUnsigned(a, b)
End Sub

Private Sub GG(a, b, c, d, x, s, ac)
 a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
 a = RotateLeft(a, s)
 a = AddUnsigned(a, b)
End Sub

Private Sub HH(a, b, c, d, x, s, ac)
 a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
 a = RotateLeft(a, s)
 a = AddUnsigned(a, b)
End Sub

Private Sub II(a, b, c, d, x, s, ac)
 a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
 a = RotateLeft(a, s)
 a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
 Dim lMessageLength
 Dim lNumberOfWords
 Dim lWordArray()
 Dim lBytePosition
 Dim lByteCount
 Dim lWordCount
 Const MODULUS_BITS = 512
 Const CONGRUENT_BITS = 448
 lMessageLength = Len(sMessage)
 lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
 ReDim lWordArray(lNumberOfWords - 1)
 lBytePosition = 0
 lByteCount = 0
 Do Until lByteCount >= lMessageLength
  lWordCount = lByteCount \ BYTES_TO_A_WORD
  lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
  lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
  lByteCount = lByteCount + 1
 Loop
 lWordCount = lByteCount \ BYTES_TO_A_WORD
 lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
 lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
 lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
 lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
 ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
 Dim lByte
 Dim lCount
 For lCount = 0 To 3
  lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
  WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
 Next
End Function

Public Function MD5(sMessage)
 Dim x
 Dim k
 Dim AA
 Dim BB
 Dim CC
 Dim DD
 Dim a
 Dim b
 Dim c
 Dim d
 Const S11 = 7
 Const S12 = 12
 Const S13 = 17
 Const S14 = 22
 Const S21 = 5
 Const S22 = 9
 Const S23 = 14
 Const S24 = 20
 Const S31 = 4
 Const S32 = 11
 Const S33 = 16
 Const S34 = 23
 Const S41 = 6
 Const S42 = 10
 Const S43 = 15
 Const S44 = 21
 x = ConvertToWordArray(sMessage)
 a = &H67452301
 b = &HEFCDAB89
 c = &H98BADCFE
 d = &H10325476
 For k = 0 To UBound(x) Step 16
  AA = a
  BB = b
  CC = c
  DD = d
  FF a, b, c, d, x(k + 0), S11, &HD76AA478
  FF d, a, b, c, x(k + 1), S12, &HE8C7B756
  FF c, d, a, b, x(k + 2), S13, &H242070DB
  FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
  FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
  FF d, a, b, c, x(k + 5), S12, &H4787C62A
  FF c, d, a, b, x(k + 6), S13, &HA8304613
  FF b, c, d, a, x(k + 7), S14, &HFD469501
  FF a, b, c, d, x(k + 8), S11, &H698098D8
  FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
  FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
  FF b, c, d, a, x(k + 11), S14, &H895CD7BE
  FF a, b, c, d, x(k + 12), S11, &H6B901122
  FF d, a, b, c, x(k + 13), S12, &HFD987193
  FF c, d, a, b, x(k + 14), S13, &HA679438E
  FF b, c, d, a, x(k + 15), S14, &H49B40821
  GG a, b, c, d, x(k + 1), S21, &HF61E2562
  GG d, a, b, c, x(k + 6), S22, &HC040B340
  GG c, d, a, b, x(k + 11), S23, &H265E5A51
  GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
  GG a, b, c, d, x(k + 5), S21, &HD62F105D
  GG d, a, b, c, x(k + 10), S22, &H2441453
  GG c, d, a, b, x(k + 15), S23, &HD8A1E681
  GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
  GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
  GG d, a, b, c, x(k + 14), S22, &HC33707D6
  GG c, d, a, b, x(k + 3), S23, &HF4D50D87
  GG b, c, d, a, x(k + 8), S24, &H455A14ED
  GG a, b, c, d, x(k + 13), S21, &HA9E3E905
  GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
  GG c, d, a, b, x(k + 7), S23, &H676F02D9
  GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
  HH a, b, c, d, x(k + 5), S31, &HFFFA3942
  HH d, a, b, c, x(k + 8), S32, &H8771F681
  HH c, d, a, b, x(k + 11), S33, &H6D9D6122
  HH b, c, d, a, x(k + 14), S34, &HFDE5380C
  HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
  HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
  HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
  HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
  HH a, b, c, d, x(k + 13), S31, &H289B7EC6
  HH d, a, b, c, x(k + 0), S32, &HEAA127FA
  HH c, d, a, b, x(k + 3), S33, &HD4EF3085
  HH b, c, d, a, x(k + 6), S34, &H4881D05
  HH a, b, c, d, x(k + 9), S31, &HD9D4D039
  HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
  HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
  HH b, c, d, a, x(k + 2), S34, &HC4AC5665
  II a, b, c, d, x(k + 0), S41, &HF4292244
  II d, a, b, c, x(k + 7), S42, &H432AFF97
  II c, d, a, b, x(k + 14), S43, &HAB9423A7
  II b, c, d, a, x(k + 5), S44, &HFC93A039
  II a, b, c, d, x(k + 12), S41, &H655B59C3
  II d, a, b, c, x(k + 3), S42, &H8F0CCC92
  II c, d, a, b, x(k + 10), S43, &HFFEFF47D
  II b, c, d, a, x(k + 1), S44, &H85845DD1
  II a, b, c, d, x(k + 8), S41, &H6FA87E4F
  II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
  II c, d, a, b, x(k + 6), S43, &HA3014314
  II b, c, d, a, x(k + 13), S44, &H4E0811A1
  II a, b, c, d, x(k + 4), S41, &HF7537E82
  II d, a, b, c, x(k + 11), S42, &HBD3AF235
  II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
  II b, c, d, a, x(k + 9), S44, &HEB86D391
  a = AddUnsigned(a, AA)
  b = AddUnsigned(b, BB)
  c = AddUnsigned(c, CC)
  d = AddUnsigned(d, DD)
 Next
 MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function

Public Function SendPost(http, page, host, post)
 With http
  .Open "POST", "http://" & host & page, False
  .setrequestHeader "HOST", host
  .setrequestHeader "Connection", "Keep-Alive"
  .setrequestHeader "Content-Type", "application/x-www-form-urlencoded"
  .setrequestHeader "Content-Length", Len(post)
  .Send post
 End With
 SendPost = http.responseText
End Function

Public Function Response(xml As String, pass As String) As String
 Dim re As RegExp
 Dim match
 Dim auth As String
 Dim code As String
 Dim a As Integer
 
 Set re = New RegExp
 re.Pattern = "<Challenge>(\w+)</Challenge>"
 Set match = re.Execute(xml)
 If (match.count > 0) Then
  auth = match(0).SubMatches(0)
  code = ""
  xml = auth & "-" & pass
  For a = 1 To Len(xml)
   code = code & Mid(xml, a, 1) & Chr(0)
  Next
  Response = "response=" & auth & "-" & MD5(code)
 End If
End Function

Private Function parseSid(xml)
 Dim re As RegExp
 Dim match
 
 Set re = New RegExp
 re.Pattern = "<SID>(?!0{16})(\w+)</SID>"
 Set match = re.Execute(xml)
 If (match.count > 0) Then
  parseSid = match(0).SubMatches(0)
 End If
End Function

Public Function getSID() As String
    getSID = m_sid
End Function

Private Function getHTTP() As Object
  Dim m_http As Object
  
  Set m_http = Nothing
  Set m_http = CreateObject("WinHttp.WinHttpRequest.5.1")
  If m_http Is Nothing Then Set m_http = CreateObject("WinHttp.WinHttpRequest.5")
  If m_http Is Nothing Then Set m_http = CreateObject("WinHttp.WinHttpRequest")
  If m_http Is Nothing Then Set m_http = CreateObject("MSXML2.ServerXMLHTTP")
  If m_http Is Nothing Then Set m_http = CreateObject("Microsoft.XMLHTTP")
  If m_http Is Nothing Then
    MsgBox "Kein HTTP-Objekt verfügbar!", 16, "Fehler:"
    Exit Function
  Else
   Set getHTTP = m_http
  End If
End Function

Public Function Login(host As String, pass As String, user) As Boolean

    Dim auth As String
    Dim data As String
    Dim http As Object
    
    Set http = getHTTP()
    'On Error Goto 0
     m_sid = ""
     m_host = host
     m_page = "/login_sid.lua"
     http.Open "GET", "http://" & host & m_page, False
     http.Send
     auth = Response(http.responseText, pass)
     If auth <> "" Then
      If user Then
       auth = auth & "&username=" & user
      End If
      m_sid = parseSid(SendPost(http, m_page, host, auth)) ' Einloggen (lua)
     End If
     If m_sid = "" Then
      m_page = "/cgi-bin/webcm"
      data = "getpage=../html/login_sid.xml"
      http.Open "GET", "http://" & host & m_page & "?" & data, False
      http.Send
      auth = Response(http.responseText, pass)
      If auth <> "" Then
       m_sid = parseSid(SendPost(http, m_page, host, data & "&login:command/" & auth)) ' Einloggen (cgi)
      Else
       data = SendPost(http, m_page, host, "login:command/password=" & pass)
      End If
     End If
    
     Login = m_sid <> ""
     Set http = Nothing
           
End Function

Public Sub LogOut()
  Dim data As String
  Dim http As Object
  If m_sid <> "" Then
      Set http = getHTTP()
      data = SendPost(http, m_page, m_host, "security:command/logout=1&sid=" & m_sid)
      Set http = Nothing
  End If
End Sub

Public Function getCurrentPhoneBook() As String
    
  Dim bookurl As String
  Dim re As RegExp
  Dim match
  Dim http As Object
  
  Set http = getHTTP()
  
  bookurl = "http://" & m_host & "/fon_num/fonbook_list.lua?sid=" & m_sid
  ' bookurl = baseurl + MENU_URL & "&var:pagename=home&var:menu=home&" & m_sid
  
  getCurrentPhoneBook = ""
  http.Open "GET", bookurl, False
  http.Send
  If Len(http.responseText) > 0 Then
    Set re = New RegExp
    'form method="POST" action="../cgi-bin/firmwarecfg" enctype="multipart/form-data" id="export_form" name="uiPostExportForm" onsubmit="return false">
    '<input type="hidden" id="sid_export" name="sid" value="973e160e827d035f" disabled>
    '<input type="hidden" id="select_fonbook_id" name="PhonebookId" value="2" disabled>
    '<input type="hidden" id="select_fonbook_name" name="PhonebookExportName" value="TEST" disabled>
    '<input type="hidden" id="select_fonbook_export" name="PhonebookExport" value="" disabled>
    '</form>
    re.Pattern = "name=" & Chr(34) & "PhonebookId" & Chr(34) & " value=" & Chr(34) & "(\d)" & Chr(34) & ""
    Set match = re.Execute(http.responseText)
    If (match.count > 0) Then
     getCurrentPhoneBook = match(0).SubMatches(0)
    End If
  End If
  
  Set http = Nothing

End Function

Public Function createFritzPhonebookXMLFromContacts(ByRef count As Integer) As String

  Dim contactfolder As Folder
  createFritzPhonebookXMLFromContacts = "<?xml version=""1.0"" encoding=""utf-8""?>" & _
  "<phonebooks>" & _
  "<phonebook>"
  
  Set contactfolder = Session.GetDefaultFolder(olFolderContacts)
  recurseContacts contactfolder, createFritzPhonebookXMLFromContacts, count
  
  createFritzPhonebookXMLFromContacts = createFritzPhonebookXMLFromContacts & "</phonebook>" & _
  "</phonebooks>"

End Function


Private Function recurseContacts(f As Folder, ByRef xmldata As String, ByRef count As Integer)
  Dim o
  Dim ct As ContactItem
  Dim subf As Folder
  
  If InStr(1, f.Name, "Recipient Cache") > 0 Or _
    InStr(1, f.Name, "GAL Contacts") > 0 Or _
    InStr(1, f.Name, "~") > 0 Or _
    InStr(1, f.Name, "{") > 0 Then
    Exit Function
  End If
  For Each o In f.Items
    If TypeName(o) = "ContactItem" Then
      Set ct = o
      xmldata = xmldata & createFritzPhonebookXMLContact(ct, count)
    End If
  Next
  For Each subf In f.Folders
    recurseContacts subf, xmldata, count
  Next
End Function

Private Function escapeXML(s As String) As String
  
  s = Replace(s, "<", "&lt;")
  s = Replace(s, ">", "&gt;")
  s = Replace(s, "&", "&amp;")
  s = Replace(s, "'", "&apos;")
  s = Replace(s, """", "&quot;")
  escapeXML = s
  
End Function

Private Function createFritzPhonebookXMLContact(ct As ContactItem, ByRef count As Integer) As String

  Dim n As String
  Dim doexport As Boolean
  
  n = escapeXML(ct.FullName)
  If n = "" Then
    n = escapeXML(ct.CompanyName)
    If n = "" Then
      n = escapeXML(ct.NickName)
    End If
  End If
  
  doexport = ct.HomeTelephoneNumber <> "" Or ct.BusinessTelephoneNumber <> "" Or ct.MobileTelephoneNumber <> ""

  If doexport Then
    createFritzPhonebookXMLContact = "<contact>" & _
      "<category>0</category>" & _
      "<person>" & _
      "<realName>" & n & "</realName>" & _
      "</person>" & _
      "<telephony>" & _
      "<number type=""home"" prio=""1"">" & ct.HomeTelephoneNumber & " </number>" & _
      "<number type=""work"">" & ct.BusinessTelephoneNumber & "</number>" & _
      "<number type=""mobile"">" & ct.MobileTelephoneNumber & "</number>" & _
      "</telephony>" & _
      "<services>" & _
      "<email classifier=""private"">" & ct.Email1Address & "</email><email />" & _
      "</services>" & _
      "<setup>" & _
      "<ringTone />" & _
      "<ringVolume />" & _
      "</setup>" & _
      "</contact>"
      count = count + 1
  End If
End Function


' http://172.22.20.2/cgi-bin/firmwarecfg?PhoneBookImportFile=egal.xml&sid=1231451&PhonebookId=1
Public Function uploadPhoneBook(pbid As String, ByRef count As Integer) As Boolean
    
  Dim uploadurl As String
  Dim uploaddata As String
  Dim bookboundary As String
  Dim re As RegExp
  Dim match
  Dim http As Object
  
  Dim xmlcontent As String
  Dim iFile As Integer

  xmlcontent = createFritzPhonebookXMLFromContacts(count)
  Debug.Print count

  Set http = getHTTP()
  
  ' <form method="POST" action="../cgi-bin/firmwarecfg" enctype="multipart/form-data" id="uiPostImportForm" name="uiPostImportForm" onsubmit="return false">
  '<div>
  'Bitte geben Sie die Datei an, aus der die Telefonbucheinträge wieder hergestellt werden sollen. Sämtliche bisherigen Einträge im Telefonbuch "WEGA14" werden bei diesem Vorgang gelöscht.
  '<input name="sid" value="0f01426c93da8b44" type="hidden">
  '<input name="PhonebookId" value="1" type="hidden">
  '<div class="formular">
  '<br>
  '<input size="40" value="" name="PhonebookImportFile" id="PhonebookImportFile" type="file">
  '</div>
  '</div>
  '</form>

  bookboundary = "-----------bookboundary"
  uploaddata = ""
  uploaddata = uploaddata & "--" & bookboundary & vbCrLf
  uploaddata = uploaddata & "Content-Disposition: form-data; name=" & Chr(34) & "sid" & Chr(34) & vbCrLf
  uploaddata = uploaddata & vbCrLf
  uploaddata = uploaddata & m_sid & vbCrLf
  uploaddata = uploaddata & "--" & bookboundary & vbCrLf
  uploaddata = uploaddata & "Content-Disposition: form-data; name=" & Chr(34) & "PhonebookId" & Chr(34) & vbCrLf
  uploaddata = uploaddata & vbCrLf
  uploaddata = uploaddata & pbid & vbCrLf
  uploaddata = uploaddata & "--" & bookboundary & vbCrLf
  uploaddata = uploaddata & "Content-Disposition: form-data; name=" & Chr(34) & "PhonebookImportFile" & Chr(34) & "; filename=" & Chr(34) & "FritzExport.xml" & Chr(34) & vbCrLf
  uploaddata = uploaddata & "Content-Type: text/xml" & vbCrLf
  uploaddata = uploaddata & vbCrLf
  uploaddata = uploaddata & xmlcontent & vbCrLf
  uploaddata = uploaddata & "--" & bookboundary & vbCrLf
    
  uploadurl = "http://" & m_host & "/cgi-bin/firmwarecfg"
  
  uploadPhoneBook = False
  http.Open "POST", uploadurl, False
  http.setrequestHeader "Content-Type", "multipart/form-data; boundary=" & bookboundary
  http.Send uploaddata
  If Len(http.responseText) > 0 Then
     If InStr(1, http.responseText, "FRITZ!Box Telefonbuch wiederhergestellt") > 1 Then
      uploadPhoneBook = True
     End If
    'Debug.Print http.responseText
  End If
  
  Set http = Nothing

End Function

