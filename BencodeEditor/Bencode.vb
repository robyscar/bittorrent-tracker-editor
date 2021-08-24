' http://demon.tw/my-work/vbs-bencode.html

'Author: Demon
'Website: http://demon.tw
'Date: 2011/4/24
Function decode_int(x, f)
    f = f + 1
    Dim newf : newf = InStr(f, x, "e")
    Dim n : n = CLng(Mid(x, f, newf-f))
    If Mid(x, f, 1) = "-" And Mid(x, f+1, 1) = "0" Then
        Err.Raise 7575, ,"ValueError"
    ElseIf Mid(x, f, 1) = "0" And newf <> f+1 Then
        Err.Raise 7575, ,"ValueError"
    End If
    decode_int = Array(n, newf+1)
End Function

Function decode_string(x, f)
    Dim colon : colon = InStr(f, x, ":")
    Dim n : n = CLng(Mid(x, f, colon-f))
    If Mid(x, f, 1) = "0" And colon <> f+1 Then
        Err.Raise 7575, ,"ValueError"
    End If
    colon = colon + 1
    decode_string = Array(Mid(x, colon, n), colon+n)
End Function

Function decode_list(x, f)
    f = f + 1
    Dim r(), count
    While Mid(x, f, 1) <> "e"
        Dim c : c = Mid(x, f, 1)
        Dim a
        Select Case c
            Case "l"
                a = decode_list(x, f)
            Case "d"
                a = decode_dict(x, f)
            Case "i"
                a = decode_int(x, f)
            Case "0","1","2","3","4","5","6","7","8","9"
                a = decode_string(x, f)
        End Select

        ReDim Preserve r(count)
        If TypeName(a(0)) = "Dictionary" Then
            Set r(count) = a(0)
        Else
            r(count) = a(0)
        End If
        f = a(1)
        count = count + 1
    Wend
    decode_list = Array(r, f+1)
End Function

Function decode_dict(x, f)
    f = f + 1
    Dim r : Set r = CreateObject("scripting.dictionary")
    While Mid(x, f, 1) <> "e"
        Dim a : a = decode_string(x, f)
        Dim k : k = a(0) : f = a(1)
        Dim c : c = Mid(x, f, 1)
        Select Case c
            Case "l"
                a = decode_list(x, f)
            Case "d"
                a = decode_dict(x, f)
            Case "i"
                a = decode_int(x, f)
            Case "0","1","2","3","4","5","6","7","8","9"
                a = decode_string(x, f)
        End Select
        If TypeName(a(0)) = "Dictionary" Then
            Set r(k) = a(0)
        Else
            r(k) = a(0)
        End If
        f = a(1)
    Wend
    decode_dict = Array(r, f+1)
End Function

' x is a string containing bencoded data, 
' where each charCodeAt value matches the byte of data
Function bdecode(x)
    'On Error Resume Next
    Dim c : c = Mid(x, 1, 1)
    Dim a
    Select Case c
        Case "l"
            a = decode_list(x, 1)
        Case "d"
            a = decode_dict(x, 1)
        Case "i"
            a = decode_int(x, 1)
        Case "0","1","2","3","4","5","6","7","8","9"
            a = decode_string(x, 1)
    End Select
    Dim r
    If TypeName(a(0)) = "Dictionary" Then
        Set r = a(0)
    Else
        r = a(0)
    End If
    Dim l : l = a(1)
    If Err.Number <> 0 Then
        Err.Raise 8732, ,"not a valid bencoded string"
    End If
    If l <> Len(str) + 1 Then
        Err.Raise 8732, ,"not a valid bencoded string"
    End If
    Set bdecode = r
End Function

'Author: Demon
'Website: http://demon.tw
'Date: 2011/4/24

Function encode_int(x, ByRef r)
    Dim n : n = UBound(r)
    ReDim Preserve r(n+3)
    r(n+1) = "i" : r(n+2) = x & "" : r(n+3) = "e"
End Function

Function encode_string(x, ByRef r)
    Dim n : n = UBound(r)
    ReDim Preserve r(n+3)
    r(n+1) = Len(x) & "" : r(n+2) = ":" : r(n+3) = x
End Function

Function encode_list(x, ByRef r)
    Dim n : n = UBound(r)
    ReDim Preserve r(n+1)
    r(n+1) = "l"
    For Each i In x
        Dim t : t = TypeName(i)
        Select Case t
            Case "Integer","Long"
                Call encode_int(i, r)
            Case "String"
                Call encode_string(i, r)
            Case "Variant()"
                Call encode_list(i, r)
            Case "Dictionary"
                Call encode_dict(i, r)
        End Select
    Next
    n = UBound(r)
    ReDim Preserve r(n+1)
    r(n+1) = "e"
End Function

Function encode_dict(x, ByRef r)
    Dim n : n = UBound(r)
    ReDim Preserve r(n+1)
    r(n+1) = "d"
    Dim keys : keys = x.Keys
    Dim length : length = UBound(keys)
    For i = 0 To length - 1
        For j = i To length
            If StrComp(keys(i), keys(j), vbTextCompare) > 0 Then
                Dim tmp
                tmp = keys(i) : keys(i) = keys(j) : keys(j) = tmp
            End If
        Next
    Next
    Dim ilist : Set ilist = CreateObject("scripting.dictionary")
    For Each i In Keys
        If TypeName(x(i)) = "Dictionary" Then
            Set ilist(i) = x(i)
        Else
            ilist(i) = x(i)
        End If
    Next
    For Each k In ilist
        n = UBound(r)
        ReDim Preserve r(n+3)
        r(n+1) = Len(k) & "" : r(n+2) = ":" : r(n+3) = k
        Dim v
        If TypeName(x(k)) = "Dictionary" Then
            Set v = x(k)
        Else
            v = x(k)
        End If
        Dim t : t = TypeName(v)
        Select Case t
            Case "Integer","Long"
                Call encode_int(v, r)
            Case "String"
                Call encode_string(v, r)
            Case "Variant()"
                Call encode_list(v, r)
            Case "Dictionary"
                Call encode_dict(v, r)
        End Select
    Next
    n = UBound(r)
    ReDim Preserve r(n+1)
    r(n+1) = "e"
End Function

Function bencode(x)
    Dim r() : ReDim r(0)
    Dim t : t = TypeName(x)
    Select Case t
        Case "Integer","Long"
            Call encode_int(x, r)
        Case "String"
            Call encode_string(x, r)
        Case "Variant()"
            Call encode_list(x, r)
        Case "Dictionary"
            Call encode_dict(x, r)
    End Select
    bencode = Join(r, "")
End Function
VBS中变量赋值还要区分对象变量和普通变量，对象变量的赋值还要多加一个Set，真是太蛋疼了，越发的觉得VBS没有JS好用。

下面简单的演示下用法：

Function read(path)
    Dim cp1252Chars : cp1252Chars = Array("\u20AC","\u201A","\u0192","\u201E","\u2026","\u2020","\u2021","\u02C6","\u2030","\u0160","\u2039","\u0152","\u017D","\u2018","\u2019","\u201C","\u201D","\u2022","\u2013","\u2014","\u02DC","\u2122","\u0161","\u203A","\u0153","\u017E","\u0178")
    Dim latin1Chars : latin1Chars = Array(ChrW("&H0080"),ChrW("&H0082"),ChrW("&H0083"),ChrW("&H0084"),ChrW("&H0085"),ChrW("&H0086"),ChrW("&H0087"),ChrW("&H0088"),ChrW("&H0089"),ChrW("&H008A"),ChrW("&H008B"),ChrW("&H008C"),ChrW("&H008E"),ChrW("&H0091"),ChrW("&H0092"),ChrW("&H0093"),ChrW("&H0094"),ChrW("&H0095"),ChrW("&H0096"),ChrW("&H0097"),ChrW("&H0098"),ChrW("&H0099"),ChrW("&H009A"),ChrW("&H009B"),ChrW("&H009C"),ChrW("&H009E"),ChrW("&H009F"))
    Dim ado : Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2 : ado.Charset = "iso-8859-1" : ado.Open
    ado.LoadFromFile path
    Dim s : s = ado.ReadText
    Dim regex : Set regex = New RegExp
    regex.Global = True
    For i = 0 To 26
        regex.Pattern = cp1252Chars(i)
        s = regex.Replace(s, latin1Chars(i))
    Next
    read = s
End Function

Function write(data, path)
    Dim ado : Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2 : ado.Charset = "iso-8859-1" : ado.Open
    ado.WriteText data
    ado.SaveToFile path, 2
End Function

str = read("foo.torrent")
' use "Set" because bdecode return a dictionary object
Set dic = bdecode(str)
' get the announce url of the tracker
announce = dic("announce");
' get the name of the torrent
name = dic("info")("name");
' get the number of files of the torrent (assuming a multi-file torrent)
number = dic("info")("files").length;
' get the size of the first file of the torrent (assuming a multi-file torrent)
number = dic("info")("files")(0)("length");
' change the announce url
dic("announce") = "http://demon.tw";
' and then encode it back to string
new_str = bencode(dic);
' then write it back to a torrent file 
' now the torrent's announce url has been changed to "http://demon.tw"
write(new_str, "bar.torrent");
