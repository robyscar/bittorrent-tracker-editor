' http://demon.tw/my-work/vbs-json.html

' VbsJson.vbs

Class VbsJson
    'Author: Demon
    'Date: 2012/5/3
    'Website: http://demon.tw
    Private Whitespace, NumberRegex, StringChunk
    Private b, f, r, n, t

    Private Sub Class_Initialize
        Whitespace = " " & vbTab & vbCr & vbLf
        b = ChrW(8)
        f = vbFormFeed
        r = vbCr
        n = vbLf
        t = vbTab

        Set NumberRegex = New RegExp
        NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
        NumberRegex.Global = False
        NumberRegex.MultiLine = True
        NumberRegex.IgnoreCase = True

        Set StringChunk = New RegExp
        StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
        StringChunk.Global = False
        StringChunk.MultiLine = True
        StringChunk.IgnoreCase = True
    End Sub
    
    'Return a JSON string representation of a VBScript data structure
    'Supports the following objects and types
    '+-------------------+---------------+
    '| VBScript          | JSON          |
    '+===================+===============+
    '| Dictionary        | object        |
    '+-------------------+---------------+
    '| Array             | array         |
    '+-------------------+---------------+
    '| String            | string        |
    '+-------------------+---------------+
    '| Number            | number        |
    '+-------------------+---------------+
    '| True              | true          |
    '+-------------------+---------------+
    '| False             | false         |
    '+-------------------+---------------+
    '| Null              | null          |
    '+-------------------+---------------+
    Public Function Encode(ByRef obj)
        Dim buf, i, c, g
        Set buf = CreateObject("Scripting.Dictionary")
        Select Case VarType(obj)
            Case vbNull
                buf.Add buf.Count, "null"
            Case vbBoolean
                If obj Then
                    buf.Add buf.Count, "true"
                Else
                    buf.Add buf.Count, "false"
                End If
            Case vbInteger, vbLong, vbSingle, vbDouble
                buf.Add buf.Count, obj
            Case vbString
                buf.Add buf.Count, """"
                For i = 1 To Len(obj)
                    c = Mid(obj, i, 1)
                    Select Case c
                        Case """" buf.Add buf.Count, "\"""
                        Case "\"  buf.Add buf.Count, "\\"
                        Case "/"  buf.Add buf.Count, "/"
                        Case b    buf.Add buf.Count, "\b"
                        Case f    buf.Add buf.Count, "\f"
                        Case r    buf.Add buf.Count, "\r"
                        Case n    buf.Add buf.Count, "\n"
                        Case t    buf.Add buf.Count, "\t"
                        Case Else
                            If AscW(c) >= 0 And AscW(c) <= 31 Then
                                c = Right("0" & Hex(AscW(c)), 2)
                                buf.Add buf.Count, "\u00" & c
                            Else
                                buf.Add buf.Count, c
                            End If
                    End Select
                Next
                buf.Add buf.Count, """"
            Case vbArray + vbVariant
                g = True
                buf.Add buf.Count, "["
                For Each i In obj
                    If g Then g = False Else buf.Add buf.Count, ","
                    buf.Add buf.Count, Encode(i)
                Next
                buf.Add buf.Count, "]"
            Case vbObject
                If TypeName(obj) = "Dictionary" Then
                    g = True
                    buf.Add buf.Count, "{"
                    For Each i In obj
                        If g Then g = False Else buf.Add buf.Count, ","
                        buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
                    Next
                    buf.Add buf.Count, "}"
                Else
                    Err.Raise 8732,,"None dictionary object"
                End If
            Case Else
                buf.Add buf.Count, """" & CStr(obj) & """"
        End Select
        Encode = Join(buf.Items, "")
    End Function

    'Return the VBScript representation of ``str(``
    'Performs the following translations in decoding
    '+---------------+-------------------+
    '| JSON          | VBScript          |
    '+===============+===================+
    '| object        | Dictionary        |
    '+---------------+-------------------+
    '| array         | Array             |
    '+---------------+-------------------+
    '| string        | String            |
    '+---------------+-------------------+
    '| number        | Double            |
    '+---------------+-------------------+
    '| true          | True              |
    '+---------------+-------------------+
    '| false         | False             |
    '+---------------+-------------------+
    '| null          | Null              |
    '+---------------+-------------------+
    Public Function Decode(ByRef str)
        Dim idx
        idx = SkipWhitespace(str, 1)

        If Mid(str, idx, 1) = "{" Then
            Set Decode = ScanOnce(str, 1)
        Else
            Decode = ScanOnce(str, 1)
        End If
    End Function
    
    Private Function ScanOnce(ByRef str, ByRef idx)
        Dim c, ms

        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "{" Then
            idx = idx + 1
            Set ScanOnce = ParseObject(str, idx)
            Exit Function
        ElseIf c = "[" Then
            idx = idx + 1
            ScanOnce = ParseArray(str, idx)
            Exit Function
        ElseIf c = """" Then
            idx = idx + 1
            ScanOnce = ParseString(str, idx)
            Exit Function
        ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = Null
            Exit Function
        ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
            idx = idx + 4
            ScanOnce = True
            Exit Function
        ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
            idx = idx + 5
            ScanOnce = False
            Exit Function
        End If
        
        Set ms = NumberRegex.Execute(Mid(str, idx))
        If ms.Count = 1 Then
            idx = idx + ms(0).Length
            ScanOnce = CDbl(ms(0))
            Exit Function
        End If
        
        Err.Raise 8732,,"No JSON object could be ScanOnced"
    End Function

    Private Function ParseObject(ByRef str, ByRef idx)
        Dim c, key, value
        Set ParseObject = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)
        
        If c = "}" Then
            Exit Function
        ElseIf c <> """" Then
            Err.Raise 8732,,"Expecting property name"
        End If

        idx = idx + 1
        
        Do
            key = ParseString(str, idx)

            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) <> ":" Then
                Err.Raise 8732,,"Expecting : delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            ParseObject.Add key, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "}" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = SkipWhitespace(str, idx + 1)
            c = Mid(str, idx, 1)
            If c <> """" Then
                Err.Raise 8732,,"Expecting property name"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
    End Function
    
    Private Function ParseArray(ByRef str, ByRef idx)
        Dim c, values, value
        Set values = CreateObject("Scripting.Dictionary")
        idx = SkipWhitespace(str, idx)
        c = Mid(str, idx, 1)

        If c = "]" Then
            ParseArray = values.Items
            Exit Function
        End If

        Do
            idx = SkipWhitespace(str, idx)
            If Mid(str, idx, 1) = "{" Then
                Set value = ScanOnce(str, idx)
            Else
                value = ScanOnce(str, idx)
            End If
            values.Add values.Count, value

            idx = SkipWhitespace(str, idx)
            c = Mid(str, idx, 1)
            If c = "]" Then
                Exit Do
            ElseIf c <> "," Then
                Err.Raise 8732,,"Expecting , delimiter"
            End If

            idx = idx + 1
        Loop

        idx = idx + 1
        ParseArray = values.Items
    End Function
    
    Private Function ParseString(ByRef str, ByRef idx)
        Dim chunks, content, terminator, ms, esc, char
        Set chunks = CreateObject("Scripting.Dictionary")

        Do
            Set ms = StringChunk.Execute(Mid(str, idx))
            If ms.Count = 0 Then
                Err.Raise 8732,,"Unterminated string starting"
            End If
            
            content = ms(0).Submatches(0)
            terminator = ms(0).Submatches(1)
            If Len(content) > 0 Then
                chunks.Add chunks.Count, content
            End If
            
            idx = idx + ms(0).Length
            
            If terminator = """" Then
                Exit Do
            ElseIf terminator <> "\" Then
                Err.Raise 8732,,"Invalid control character"
            End If
            
            esc = Mid(str, idx, 1)

            If esc <> "u" Then
                Select Case esc
                    Case """" char = """"
                    Case "\"  char = "\"
                    Case "/"  char = "/"
                    Case "b"  char = b
                    Case "f"  char = f
                    Case "n"  char = n
                    Case "r"  char = r
                    Case "t"  char = t
                    Case Else Err.Raise 8732,,"Invalid escape"
                End Select
                idx = idx + 1
            Else
                char = ChrW("&H" & Mid(str, idx + 1, 4))
                idx = idx + 5
            End If

            chunks.Add chunks.Count, char
        Loop

        ParseString = Join(chunks.Items, "")
    End Function

    Private Function SkipWhitespace(ByRef str, ByVal idx)
        Do While idx <= Len(str) And _
            InStr(Whitespace, Mid(str, idx, 1)) > 0
            idx = idx + 1
        Loop
        SkipWhitespace = idx
    End Function


Function ParseJson(strJson)
    Set html = CreateObject("htmlfile")
    Set window = html.parentWindow
    window.execScript "var json = " & strJson, "JScript"
    Set ParseJson = window.json
End Function


End Class

'''     Example.vbs
'''     
'''     'Author: Demon
'''     'Date: 2012/5/3
'''     'Website: http://demon.tw
'''     Dim fso, json, str, o, i
'''     Set json = New VbsJson
'''     Set fso = WScript.CreateObject("Scripting.Filesystemobject")
'''     str = fso.OpenTextFile("json.txt").ReadAll
'''     Set o = json.Decode(str)
'''     WScript.Echo o("Image")("Width")
'''     WScript.Echo o("Image")("Height")
'''     WScript.Echo o("Image")("Title")
'''     WScript.Echo o("Image")("Thumbnail")("Url")
'''     For Each i In o("Image")("IDs")
'''         WScript.Echo i
'''     Next
'''     

'''     json.txt
'''     
'''     {
'''        "Image": {
'''            "Width":  800,
'''            "Height": 600,
'''            "Title":  "View from 15th Floor",
'''            "Thumbnail": {
'''                "Url":    "http://www.example.com/image/481989943",
'''                "Height": 125,
'''                "Width":  "100"
'''            },
'''            "IDs": [116, 943, 234, 38793]
'''          }
'''     }