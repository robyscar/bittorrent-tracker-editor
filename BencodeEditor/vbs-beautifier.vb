' http://demon.tw/my-work/vbs-beautifier.html

Option Explicit

If WScript.Arguments.Count = 0 Then 
    MsgBox "Please drag the code file to be formatted onto this file" , vbInformation , "How to use" 
    WScript.Quit 
End If

'Author: Demon'Time 
: 2011/12/24'Link 
: http://demon.tw/my-work/vbs-beautifier.html'Description 
: VBScript code formatting tool 
' Note: 
'1. Wrong VBScript code Cannot be formatted correctly'2. 
The code cannot contain template tags such as %[comment]% %[quoted]%, etc., which needs to be improved 
' 3. As can be seen from 2, the tool cannot format itself

Dim Beautifier, i
Set Beautifier = New VbsBeautifier

For Each i In WScript.Arguments
    Beautifier.BeautifyFile i
Next

MsgBox "Code formatting completed" , vbInformation , "Prompt"


Class VbsBeautifier
    'VbsBeautifier类

    Private quoted, comments, code, indents
    Private ReservedWord, BuiltInFunction, BuiltInConstants, VersionInfo

    'Public method 
    ' format string 
    Public Function Beautify( ByVal input) 
        code = input 
        code = Replace (code, vbCrLf , vbLf )

        Call GetQuoted()
        Call GetComments()
        Call GetErrorHandling()

        Call ColonToNewLine()
        Call FixSpaces()
        Call ReplaceReservedWord()
        Call InsertIndent()
        Call FixIndent()

        Call PutErrorHandling()
        Call PutComments()
        Call PutQuoted()

        code = Replace(code, vbLf, vbCrLf)
        code =  VersionInfo & code
        Beautify = code
    End Function

    'Public method 
    ' format file 
    Public Function BeautifyFile( ByVal path) 
        Dim fso 
        Set fso = CreateObject ( "scripting.filesystemobject" ) 
        BeautifyFile = Beautify(fso.OpenTextFile(path).ReadAll) 
        'Back up the file to avoid errors 
        fso.GetFile(path ).Copy path & ".bak" , True 
        fso.OpenTextFile(path, 2 , True ).Write(BeautifyFile) 
    End Function

    Private Sub Class_Initialize()
        '保留字
        ReservedWord = "And As Boolean ByRef Byte ByVal Call Case Class Const Currency Debug Dim Do Double Each Else ElseIf Empty End EndIf Enum Eqv Event Exit Explicit False For Function Get Goto If Imp Implements In Integer Is Let Like Long Loop LSet Me Mod New Next Not Nothing Null On Option Optional Or ParamArray Preserve Private Property Public RaiseEvent ReDim Rem Resume RSet Select Set Shared Single Static Stop Sub Then To True Type TypeOf Until Variant WEnd While With Xor"
        '内置函数
        BuiltInFunction = "Abs Array Asc Atn CBool CByte CCur CDate CDbl CInt CLng CSng CStr Chr Cos CreateObject Date DateAdd DateDiff DatePart DateSerial DateValue Day Escape Eval Exp Filter Fix FormatCurrency FormatDateTime FormatNumber FormatPercent GetLocale GetObject GetRef Hex Hour InStr InStrRev InputBox Int IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase LTrim Left Len LoadPicture Log Mid Minute Month MonthName MsgBox Now Oct Randomize RGB RTrim Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second SetLocale Sgn Sin Space Split Sqr StrComp StrReverse String Tan Time TimeSerial TimeValue Timer Trim TypeName UBound UCase Unescape VarType Weekday WeekdayName Year"
        '内置常量
        BuiltInConstants = "VbBlack vbRed vbGreen vbYellow vbBlue vbMagenta vbCyan vbWhite vbBinaryCompare vbTextCompare vbSunday vbMonday vbTuesday vbWednesday vbThursday vbFriday vbSaturday vbUseSystemDayOfWeek vbFirstJan1 vbFirstFourDays vbFirstFullWeek vbGeneralDate vbLongDate vbShortDate vbLongTime vbShortTime vbObjectError vbOKOnly vbOKCancel vbAbortRetryIgnore vbYesNoCancel vbYesNo vbRetryCancel vbCritical vbQuestion vbExclamation vbInformation vbDefaultButton1 vbDefaultButton2 vbDefaultButton3 vbDefaultButton4 vbApplicationModal vbSystemModal vbOK vbCancel vbAbort vbRetry vbIgnore or similar vbNo vbCr vbCrLf vbFormFeed vbLf vbNewLine vbNullChar vbNullString vbTab vbVerticalTab vbUseDefault vbTrue vbFalse vbEmpty vbNull vbInteger vbLong vbSingle vbDouble vbCurrency vbDate vbString vbObject vbError vbBoolean vbVariant vbDataObject vbDecimal vbByte VBArray WScript "
        '版本信息
        VersionInfo = Chr(39) & Chr(86) & Chr(98) & Chr(115) & Chr(66) & Chr(101) & Chr(97) & Chr(117) & Chr(116) & Chr(105) & Chr(102) & Chr(105) & Chr(101) & Chr(114) & Chr(32) & Chr(49) & Chr(46) & Chr(48) & Chr(32) & Chr(98) & Chr(121) & Chr(32) & Chr(68) & Chr(101) & Chr(109) & Chr(111) & Chr(110) & Chr(13) & Chr(10) & Chr(39) & Chr(104) & Chr(116) & Chr(116) & Chr(112) & Chr(58) & Chr(47) & Chr(47) & Chr(100) & Chr(101) & Chr(109) & Chr(111) & Chr(110) & Chr(46) & Chr(116) & Chr(119) & Chr(13) & Chr(10)
        '缩进大小
        Set indents = CreateObject("scripting.dictionary")
        indents("if") = 1
        indents("sub") = 1
        indents("function") = 1
        indents("property") = 1
        indents("for") = 1
        indents("while") = 1
        indents("do") = 1
        indents("for") = 1
        indents("select") = 1
        indents("with") = 1
        indents("class") = 1
        indents("end") = -1
        indents("next") = -1
        indents("loop") = -1
        indents("wend") = -1
    End Sub

    Private Sub Class_Terminate() 
        'Do nothing 
    End Sub

    '将字符串替换成%[quoted]%
    Private Sub GetQuoted()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.Pattern = """.*?"""
        Set quoted = re.Execute(code)
        code = re.Replace(code, "%[quoted]%")
    End Sub

    '将%[quoted]%替换回字符串
    Private Sub PutQuoted()
        Dim i
        For Each i In quoted
            code = Replace(code, "%[quoted]%", i, 1, 1)
        Next
    End Sub

    '将注释替换成%[comment]%
    Private Sub GetComments()
        Dim re
        Set re = New RegExp
        re.Global = True
        re.Pattern = "'.*"
        Set comments = re.Execute(code)
        code = re.Replace(code, "%[comment]%")
    End Sub

    'Replace %[comment]% back to the comment 
    Private Sub PutComments() 
        Dim i 
        For Each i In comments 
            code = Replace (code, "%[comment]%" , i, 1 , 1 ) 
        Next 
    End Sub

    'Replace the colon with a newline 
    Private Sub ColonToNewLine 
        code = Replace (code, ":" , vbLf ) 
    End Sub

    'Replace the error handling statement with the template tag 
    Private Sub GetErrorHandling() 
        Dim re 
        Set re = New RegExp 
        re.Global = True 
        re.IgnoreCase = True 
        re.Pattern = "on\s+error\s+resume\s+next" 
        code = re. Replace (code, "%[resumenext]%" ) 
        re.Pattern = "on\s+error\s+goto\s+0" 
        code = re. Replace (code, "%[gotozero]%" ) 
    End Sub

    'Replace the template tag back to the error handling statement 
    Private Sub PutErrorHandling() 
        code = Replace (code, "%[resumenext]%" , "On Error Resume Next" ) 
        code = Replace (code, "%[gotozero]%" , " On Error GoTo 0" ) 
    End Sub

    'Formatted spaces 
    Private Sub FixSpaces () 
        Dim Re 
        the Set Re = New the RegExp 
        re.Global = True 
        re.IgnoreCase = True 
        re.MultiLine = True 
        ' strips spaces before and after each row 
        re.Pattern = "^ [\ T] * ( .*?)[ \t]*$" 
        code = re. Replace (code, "$1" ) 
        'Add spaces before and after the operator 
        re.Pattern = "[ \t]*(=|<|>|-|\ +|&|\*|/|\^|\\)[ \t]*" 
        code = re. Replace (code, "$1" ) 
        'Remove the space between <> 
        re.Pattern = "[ \t]* <\s*>[ \t]*" 
        code = re. Replace (code," <> ")
        'Remove the <= space in the middle 
        re.Pattern = "[ \t]*<\s*=[ \t]*" 
        code = re. Replace (code, "<=" ) 
        'Remove>= the space in the middle 
        re. Pattern = "[ \t]*>\s*=[ \t]*" 
        code = re. Replace (code, ">=" ) 
        'Add a space before the _ at the end of the line 
        re.Pattern = "[ \t ]*_[ \t]*$" 
        code = re. Replace (code, "_" ) 
        'Remove the extra spaces in the Do While 
        re.Pattern = "[ \t]*Do\s*While[ \t]* " 
        code = re. Replace (code, "Do While" ) 
        'Remove the extra space between Do Until 
        re.Pattern = "[ \t]*Do\s*Until[ \t]*"
        code = re.Replace (code, "Do Until" ) 
        'Remove the extra spaces in the End Sub 
        re.Pattern = "[ \t]*End\s*Sub[ \t]*" 
        code = re. Replace (code, "End Sub" ) 
        'Remove the extra spaces in the End Function 
        re.Pattern = "[ \t]*End\s*Function[ \t]*" 
        code = re. Replace (code, "End Function" ) 
        'Remove the 
        extra spaces in the End If Space re.Pattern = "[ \t]*End\s*If[ \t]*" 
        code = re. Replace (code, "End If" ) 
        'Remove the extra space between End With 
        re.Pattern = "[ \ t]*End\s*With[ \t]*" 
        code = re.Replace(code, "End With")
        'Remove the extra spaces in the End Select 
        re.Pattern = "[ \t]*End\s*Select[ \t]*" 
        code = re. Replace (code, "End Select" ) 
        'Remove the extra spaces in the Select Case 
        re.Pattern = "[ \t]*Select\s*Case[ \t]*" 
        code = re. Replace (code, "Select Case " ) 
    End Sub

    'Replace the reserved word built-in function built-in constants with initial capitalization 
    Private Sub ReplaceReservedWord() 
        Dim re, words, word 
        Set re = New RegExp 
        re.Global = True 
        re.IgnoreCase = True 
        re.MultiLine = True

        words = Split(ReservedWord, " ")
        For Each word In words
            re.Pattern = "(\b)" & word & "(\b)"
            code = re.Replace(code, "$1" & word & "$2")
        Next

        words = Split(BuiltInFunction, " ")
        For Each word In words
            re.Pattern = "(\b)" & word & "(\b)"
            code = re.Replace(code, "$1" & word & "$2")
        Next

        words = Split(BuiltInConstants, " ")
        For Each word In words
            re.Pattern = "(\b)" & word & "(\b)"
            code = re.Replace(code, "$1" & word & "$2")
        Next
    End Sub

    '插入缩进
    Private Sub InsertIndent()
        Dim lines, line, i, n, t, delta
        lines = Split(code, vbLf)
        n = UBound(lines)
        For i = 0 To n
            line = lines(i)
            SingleLineIfThen line
            t = delta
            delta = delta + CountDelta(line)

            If t <= delta Then
                lines(i) = String(t, vbTab) & lines(i)
            Else
                lines(i) = String(delta, vbTab) & lines(i)
            End If
        Next
        code = Join(lines, vbLf)
    End Sub

    '调整错误的缩进
    Private Sub FixIndent()
        Dim lines, i, n, re
        Set re = New RegExp
        re.IgnoreCase = True
        lines = Split(code, vbLf)
        n = UBound(lines)
        For i = 0 To n
            re.Pattern = "^\t*else"
            If re.Test(lines(i)) Then
                lines(i) = Replace(lines(i), vbTab, "", 1, 1)
            End If
        Next
        code = Join(lines, vbLf)
    End Sub

    '计算缩进大小
    Private Function CountDelta(ByRef line)
        Dim i, re, delta
        Set re = New RegExp
        re.Global = True
        re.IgnoreCase = True
        For Each i In indents.Keys
            re.Pattern = "^\s*\b" & i & "\b"
            If re.Test(line) Then
                '方便调试
                'WScript.Echo line
                line = re.Replace(line, "")
                delta = delta + indents(i)
            End If
        Next
        CountDelta = delta
    End Function

    '处理单行的If Then
    Private Sub SingleLineIfThen(ByRef line)
        Dim re
        Set re = New RegExp
        re.IgnoreCase = True
        re.Pattern = "if.*?then.+"
        line = re.Replace(line, "")
        '去掉Private Public前缀
        re.Pattern = "(private|public).+?(sub|function|property)"
        line = re.Replace(line, "$2")
    End Sub

End 
Class'Demon, on Christmas Eve 2011