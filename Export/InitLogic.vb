'=================================================================================
' 开发者：猎人幻想
'=================================================================================
#If Win64 Then
    Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
    Public Declare Function GetTickCount Lib "kernel32" () As Long
#End If
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public msgDistPath As String, msgServerPath As String, msgClientPath As String, enMsgClientPath As String
Public messageDict As New Dictionary
Public englishDict As New Dictionary
Public localPath As String
Public MsgCount As Long '记录所有"MSG:"个数
Public msgBases As New Dictionary '记录所有"MSG:"内容

'定义一个子程序
Function BrowDir() As String
    Dim bi As BROWSEINFO
    Dim pidl&, rtn&, path$, pos%
    pidl& = SHBrowseForFolder(bi)
    path$ = Space$(512)
    rtn& = SHGetPathFromIDList(ByVal pidl&, ByVal path$)
    If rtn& Then
    pos% = InStr(path$, Chr$(0))
    BrowDir = Left(path$, pos - 1)
    If Right(BrowDir, 1) <> "\" And Trim(BrowDir) <> "" Then
    BrowDir = BrowDir + "\"
    End If
    End If '执行Browdir子程序即可弹出目录选择窗口了,过程中MyPath返回已选择的路径字串.
End Function

' all   全表导出
' quick 快速导出
Sub OnExport(ByVal all As Boolean, ByVal quick As Boolean)
    On Error Resume Next
    Dim st As Worksheet
    msgClientPath = ""
    enMsgClientPath = ""
    localPath = Workbooks("导出工具2.0.xlam").path
    Dim client As String
    Dim server As String
    Dim Distribute As String
    Dim xXport As New DOMDocument60
    xXport.Load localPath & "\xport.xml"
    Call getworkbook("Settings.xls", True)

    server = GetSetting("HunterFantasy", "SetupPath", "Server", 0)
    client = GetSetting("HunterFantasy", "SetupPath", "Client", 0)
    Distribute = GetSetting("HunterFantasy", "SetupPath", "Distribute", 0)
    If server = "0" Or server = "" Or client = "0" Or Distribute = "0" Then
        MsgBox "请安装导出工具.xlam,配置关键目录后再使用此功能!"
        Exit Sub
    End If

    Call InitMessage(xXport, server, client, Distribute)
    If quick = False Then
        If all = True Then
            For Each st In ActiveWorkbook.Sheets
                Call InitMessageCount(st.Name)
                获得所有导出路径 st.Name, st.Parent.Name
            Next
        Else
            Call InitMessageCount(ActiveSheet.Name)
        End If
        Call LoadMessage
    End If
    
    If all = True Then
        For Each st In ActiveWorkbook.Sheets
            ErrorMag = new导出表格(st.Name, st.Parent.Name, xXport, server, client, Distribute)
            获得所有导出路径 st.Name, st.Parent.Name
        Next
    Else
        ErrorMsg = new导出表格(ActiveSheet.Name, ActiveWorkbook.Name, xXport, server, client, Distribute)
        If ErrorMsg <> "" Then
            MsgBox ErrorMsg
            Exit Sub
        End If
    End If
    
    If MsgCount > 0 Then
        Call SaveMessage
    ElseIf englishDict.count > 0 Then
        Call SaveEnglishMessage
    End If
    Workbooks("Settings.xls").Close
    
    If all = True Then
        If Application.DisplayAlerts Then
            If quick = True Then
                MsgBox "全表快速导出完毕"
            Else
                MsgBox "全表导出完毕"
            End If
        End If
    Else
        If quick = True Then
            MsgBox "快速导出完毕"
        Else
            MsgBox "导出完毕"
        End If
    End If
End Sub

Sub 全表导出()
    OnExport True, False
End Sub

Sub 全表快速导出()
    OnExport True, True
End Sub

Sub 导出()
    OnExport False, False
End Sub

Sub 快速导出()
    OnExport False, True
End Sub

'新版导出要求V0.9：1、导出时需打开Settings.xls
'                  2、Settings.xls中的对应表名为“表名+Title”相同
'                  3、导出表中各列的顺序与Settings.xls中对应表各行的顺序相同
Function new导出表格(ByVal 表名 As String, ByVal 薄名 As String, ByVal xXport As DOMDocument60, ByVal server As String, ByVal client As String, ByVal Distribute As String)

    On Error Resume Next
    Dim StartTime As Long
    StartTime = GetTickCount

    Dim xTemp, xMsgNode As IXMLDOMNode
    Set xTemp = xXport.SelectSingleNode("Settings/Item[@Sheet='" & 表名 & "']")
    If xTemp Is Nothing Then
        frmLog.Log 表名 & "没被登记为可导出!", False
        new导出表格 = "表单" & 表名 & "没有在XPort.xml定义。"
        Exit Function
    End If
    
    If Not xTemp.SelectSingleNode("Export") Is Nothing Then
        Dim i As Long, valiRow() As Long
        i = 0
        For Each Node In xXport.SelectNodes("Settings/Item[@Sheet='" & 表名 & "']/Export/File")
            p = Node.text
            i = i + 1
            If p <> "" Then
                tempP = p
                p = 替换(p, "[Server]", server)
                p = 替换(p, "[Client]", client)
                p = 替换(p, "[Distribute]", Distribute)
                
                valiRow = 读取导出列(i, 表名, 薄名)
                Dim isClient As Long
                isClient = InStr(p, client)
                
                
                
                ' 向客户端导.bin,向服务器导.txt
                If isClient = 0 Then
                    Call new导出TXT(p, valiRow, i, 表名, 薄名)
                Else
                    Dim path As String
                    path = p
                    path = Replace(path, ".txt", ".bin")
                    Call new导出二进制文件(path, valiRow, 表名, 薄名)
                End If
                
                '判断导出到服务器还是客户端
                Dim k As Long
                k = 0
                k = WorksheetFunction.Find("[Distribute]", tempP)
                If k = 1 Then
                    Call 导出h(表名, i, server, tempP)
                    Call 导出C(表名, i, server, tempP)
                Else
                    Call 导出CS(表名, 薄名, i, client, tempP)
                End If
            End If
        Next Node
    End If

    frmLog.Log 表名 & "导出完毕! 用时" & GetTickCount - StartTime
    new导出表格 = ""

    'Send2Window "/CPP.staticTabFileMgr:CheckAndReloadOnce()" & vbCr, "PetsClient.exe"
    Set xXport = Nothing
End Function

Sub utf8SaveAs(path, content)
    Set objStream = CreateObject("ADODB.Stream")
    Dim tempPath As Variant, tempResult As Boolean
    tempPath = 匹配文件名(path)
    Kill tempPath & ".txt"
    tempPath = tempPath & "1.txt"
    With objStream
        .Open
        .Charset = "utf-8"
        .Position = objStream.Size
        .WriteText = content
        .SaveToFile tempPath, 2
        .Close
    End With
    Set objStream = Nothing
    tempResult = Convert2utf8(tempPath, path)
End Sub

Function 匹配文件名(ByVal inputCol As String) As String
    Dim reg As New RegExp
    With reg
        .Global = True
        .IgnoreCase = True
        '.Pattern = "\d+"
        '.Pattern = "^(.+)_([0-9]+)_([0-9]+)$"
        .Pattern = "(.+)\.txt"
    End With
    Dim mc As MatchCollection, temp As String
    Set mc = reg.Execute(inputCol)

    For Each m In mc
        temp = m.SubMatches(0)
        'temp(1) = m.SubMatches(1)
    Next
    匹配文件名 = temp
End Function

Public Function Convert2utf8(ByVal fileName As String, ByVal FileTo As String) As Boolean
    Dim ReadIntFileNum, WriteIntFileNum As Long
    ReadIntFileNum = FreeFile() '获取一个空文件
    WriteIntFileNum = FreeFile() + 1
    Open fileName For Binary As #ReadIntFileNum
    Open FileTo For Binary As #WriteIntFileNum
    Dim fileByte As Long
    'Seek #ReadIntFileNum, 4
'    Get #ReadIntFileNum, , fileByte
'    Get #ReadIntFileNum, , fileByte
'    Get #ReadIntFileNum, , fileByte
Dim i
    While Not EOF(ReadIntFileNum)
        i = i + 1
        Debug.Print i
        Get #ReadIntFileNum, , fileByte
        Put #WriteIntFileNum, , fileByte
    Wend
    Close #ReadIntFileNum
    Close #WriteIntFileNum
    'Kill fileName
End Function
Function utf8Loadfile(path)
    Set objStream = CreateObject("ADODB.Stream")
    With objStream
        .Open
        .Type = 2 '设置数据类型为文本
        .Charset = "utf-8"
        .LoadFromFile path
        utf8Loadfile = .ReadText
        .Close
    End With
    Set objStream = Nothing
End Function

Function 读取导出列(ByVal index As Long, sheetName As String, bookName As String) As Variant
    Dim i, maxRow, firstCol As Long, temp As String
    
    ' 调试到这里下标越界，可能错误情形，setting 里面的Title表没有
    On Error Resume Next

    With Workbooks("Settings.xls").Sheets(sheetName & "Title")
        maxRow = .UsedRange.rows.count
        
        'sheet 不存在时，这里报错
        If maxRow = 0 Then
            MsgBox "Settings.xls里没有找到或空表单:" & sheetName & "Title", , "错误"
            Exit Function
        End If
        On Error GoTo 0
        
        Dim valRow() As Long
        ReDim valRow(maxRow - 1) As Long
        
        '找到首个*号所在列
        firstCol = 1
        Do
            firstCol = firstCol + 1
            temp = .Cells(2, firstCol)
        Loop Until temp = "*" Or temp = "key1" Or temp = "key2" Or temp = "key3"
        
        '做索引
        Dim dic As New Dictionary, tempStr As String
        For i = 2 To maxRow
            If .Cells(i, firstCol + 2 * (index - 1)) = "*" Or .Cells(i, firstCol + 2 * (index - 1)) = "key1" Or .Cells(i, firstCol + 2 * (index - 1)) = "key2" Or .Cells(i, firstCol + 2 * (index - 1)) = "key3" Then
                tempStr = .Cells(i, 1)
                dic.Add tempStr, 1
            Else
                dic.Add .Cells(i, 1), 0
            End If
        Next i
    End With
    
    '将列由字典索引到settings.xls中的对应行
    With Workbooks(bookName).Worksheets(sheetName)
        maxCol = .UsedRange.Columns.count
        Dim valiColumn() As Long
        ReDim valiColumn(maxCol - 1) As Long
        For i = 2 To maxCol
            tempStr = .Cells(1, i)
            'a = dic.item(b)
            valiColumn(i - 2) = dic.item(tempStr)
        Next
    End With
    
    读取导出列 = valiColumn

End Function

'导出.cs文件
Sub 导出CS(sheetName As String, bookName As String, index As Long, client As String, ByVal pathForTxt As String)
    ' 表头信息表
    Dim dic As New Dictionary
    '打开.CS配置文件模板
    Dim fso0 As FileSystemObject, tempFile, tempString As String, path As String
    Dim tempString1 As String, tempString2, tempString3 As String, tempArray As Variant
    
    With Workbooks("Settings.xls").Worksheets(sheetName & "Title")
        
        'Set fso0 = New FileSystemObject
        'path = localPath & "\csConfigTemplate.txt"
        
        'tempFile = utf8Loadfile(path)
        'Set tempFile = fso0.OpenTextFile(path)
        
        Dim maxRow As Long, maxColumn As Long
        maxRow = .UsedRange.rows.count
        maxColumn = .UsedRange.rows.count
       
        Dim valiRow() As Long, valiRow2() As Long
        '哪几行标了*
        ReDim valiRow(maxRow - 1) As Long
        '哪几行有自定义导出
        ReDim valiRow2(maxRow - 1) As Long

        For i = 0 To maxRow - 2
            valiRow(i) = 0
        Next
        '找到首个*号所在列
        firstCol = 1
        Do
            firstCol = firstCol + 1
            temp = .Cells(2, firstCol)
        Loop Until temp = "*" Or temp = "key1" Or temp = "key2" Or temp = "key3"
        
        Dim keyNum As Long
        keyNum = 0
        For i = 2 To maxRow
            If .Cells(i, firstCol + 2 * (index - 1)) = "*" Then
                valiRow(i - 2) = 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key1" Then
                valiRow(i - 2) = 2
                keyNum = keyNum + 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key2" Then
                valiRow(i - 2) = 3
                keyNum = keyNum + 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key3" Then
                valiRow(i - 2) = 4
                keyNum = keyNum + 1
            End If
            If .Cells(i, firstCol + 1 + 2 * (index - 1)) <> nil Then
                valiRow2(i - 2) = 1
            End If
        Next i
        
        If keyNum = 1 Then
            Call DataCheck(sheetName, bookName)
            path = localPath & "\csConfigTemplate1.txt"
        ElseIf keyNum = 2 Then
            path = localPath & "\csConfigTemplate2.txt"
        ElseIf keyNum = 3 Then
            path = localPath & "\csConfigTemplate3.txt"
        End If
        tempFile = utf8Loadfile(path)
        
        'ReDim valiRow(maxRow - 1) As Long
        
        
        tempString = tempFile
        tempString = Replace(tempString, "[[tabname]]", sheetName)
        
        Dim reg As New RegExp
        With reg
            .Global = True
            .IgnoreCase = True
            .Pattern = "\d+"
            .Pattern = "\\Config\\.*$"
        End With
        Dim mc As MatchCollection
        Set mc = reg.Execute(pathForTxt)
        pathForTxt = mc.item(0)
        pathForTxt = Replace(pathForTxt, "\", "/")
        
        pathForTxt = Replace(pathForTxt, ".txt", ".bin")
        
        tempString = Replace(tempString, "[[path]]", pathForTxt)
        
        Dim keyFlag As Long
        keyFlag = 0
        For i = 2 To maxRow
            If valiRow(i - 2) = 2 Then
                tempString = Replace(tempString, "[[key1type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key1name]]", .Cells(i, 1))
                keyFlag = 1
            ElseIf valiRow(i - 2) = 3 Then
                tempString = Replace(tempString, "[[key2type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key2name]]", .Cells(i, 1))
            ElseIf valiRow(i - 2) = 4 Then
                tempString = Replace(tempString, "[[key3type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key3name]]", .Cells(i, 1))
            End If
        Next
        
        'tempString = 全替换(tempString, "[[tabname]]", sheetname)
        'tempString = 全替换(tempString, "[[tabname]]", sheetname)
        
'        tempString1 = "public class " & sheetname & "Mgr : CCfg1KeyMgrTemplate<" & sheetname & "Mgr, int, " & sheetname & ">"
'        tempString = Replace(tempString, "public class SkillConfigInfoMgr : CCfg1KeyMgrTemplate<SkillConfigInfoMgr, int, SkillConfigInfo>", tempString1)
'
'        tempString1 = "Dictionary<int, " & sheetname & "> _un = new Dictionary<int, " & sheetname & ">();"
'        tempString = Replace(tempString, "Dictionary<int, SkillConfigInfo> _un = new Dictionary<int, SkillConfigInfo>();", tempString1)
'
'        tempString1 = "public class " & sheetname & " : ITabItemWith1Key<int>"
'        tempString = Replace(tempString, "public class SkillConfigInfo : ITabItemWith1Key<int>", tempString1)
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString1 = tempString1 & "public static readonly string __" & tempString2 & " = " & Chr(34) & tempString2 & Chr(34) & ";" & vbCrLf & vbTab & vbTab
            End If
        Next
        tempString = Replace(tempString, "//public static readonly string __[[列名]] = " & """" & "[[列名]]" & """" & ";", tempString1)
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                tempArray = 匹配数组列(tempString2)
                If tempArray(0) <> nil Then
                    If tempArray(2) = 0 Then
                        tempString1 = tempString1 & "public virtual " & tempString3 & "[] " & tempArray(0) & " { get; private set; }" & vbCrLf & vbTab & vbTab
                    End If
                Else
                    If tempString3 = "script" Then
                        tempString1 = tempString1 & "public virtual " & "int" & " " & tempString2 & " { get; private set; }" & vbCrLf & vbTab & vbTab
                    ElseIf tempString3 = "int_array" Then
                        tempString1 = tempString1 & "public virtual " & "int[]" & " " & tempString2 & " { get; private set; }" & vbCrLf & vbTab & vbTab
                    ElseIf tempString3 = "float_array" Then
                        tempString1 = tempString1 & "public virtual " & "float[]" & " " & tempString2 & " { get; private set; }" & vbCrLf & vbTab & vbTab
                    ElseIf tempString3 = "string_array" Then
                        tempString1 = tempString1 & "public virtual " & "string[]" & " " & tempString2 & " { get; private set; }" & vbCrLf & vbTab & vbTab
                    ElseIf tempString3 = "enum" Then
                        
                    Else
                        tempString1 = tempString1 & "public virtual " & tempString3 & " " & tempString2 & " { get; private set; }" & vbCrLf & vbTab & vbTab
                    End If
                End If
            End If
        Next
        tempString = Replace(tempString, "//public [[列类型]] [[列名]] { get; private set; }", tempString1)
        
        'tempString1 = "public " & sheetname & "()"
        'tempString = Replace(tempString, "public SkillConfigInfo()", tempString1)
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                tempArray = 匹配数组列(tempString2)
                If tempArray(0) <> nil Then
                    If tempArray(2) = 0 Then
                        tempString1 = tempString1 & tempArray(0) & " = new " & tempString3 & "[" & tempArray(1) & "];" & vbCrLf & vbTab & vbTab & vbTab
                    End If
                End If
            End If
        Next
        tempString = Replace(tempString, "// [[列名前缀]] = new [[列类型]][总数];", tempString1)
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                tempArray = 匹配数组列(tempString2)
                If tempArray(0) <> nil Then
                    tempString1 = tempString1 & tempArray(0) & "[" & tempArray(2) & "] " & " = tf.Get<" & tempString3 & ">(__" & tempString2 & ");" & vbCrLf & vbTab & vbTab & vbTab
                Else
                    If tempString3 = "script" Then
                        'tempString1 = tempString1 & tempString2 & " = ToLua.instance.funcBuffer.Add(tf.Get<string>(__" & tempString2 & "));" & vbCrLf & vbTab & vbTab & vbTab
                        tempString1 = tempString1 & tempString2 & " =  LuaFramework.Util.LoadString(tf.Get<string>(__" & tempString2 & "));" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "int_array" Then
                        'tempString1 = tempString1 & tempString2 & " = Array.ConvertAll<string, int>(tf.Get<string>(__" & tempString2 & ").Split(new char[] { ',' }), str => int.Parse(str));" & vbCrLf & vbTab & vbTab & vbTab
                        'tempString1 = tempString1 & "if (__" & tempString2 & " == " & """" & """" & ") " & tempString2 & " = new int[0];" & vbCrLf & vbTab & vbTab & vbTab
                        'tempString1 = tempString1 & "else " & tempString2 & " = Array.ConvertAll<string, int>(tf.Get<string>(__" & tempString2 & ").Split(new char[] { ',' }), str => int.Parse(str));" & vbCrLf & vbTab & vbTab & vbTab
                        tempString1 = tempString1 & "{string s = tf.Get<string>(__" & tempString2 & ");if (s == """") " & tempString2 & " = new int[0];else " & tempString2 & " = Array.ConvertAll<string, int>(s.Split(new char[] { ',' }), str => int.Parse(str));}" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "float_array" Then
                        tempString1 = tempString1 & "{string s = tf.Get<string>(__" & tempString2 & ");if (s == """") " & tempString2 & " = new float[0];else " & tempString2 & " = Array.ConvertAll<string, float>(s.Split(new char[] { ',' }), str => float.Parse(str));}" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "string_array" Then
                        tempString1 = tempString1 & "{string s = tf.Get<string>(__" & tempString2 & ");if (s == """") " & tempString2 & " = new string[0];else " & tempString2 & " = Array.ConvertAll<string, string>(s.Split(new char[] { ',' }), str => str);}" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "enum" Then

                    Else
                        tempString1 = tempString1 & tempString2 & " = tf.Get<" & tempString3 & ">(__" & tempString2 & ");" & vbCrLf & vbTab & vbTab & vbTab
                    End If
                End If
            End If
        Next
        tempString = Replace(tempString, "ID = tf.Get<int>(__ID);", tempString1)

        '新方法 BinaryReader 自动导出'''''''''''''''''''''''''''''''''''
        '新方法 BinaryReader 自动导出'''''''''''''''''''''''''''''''''''
        Dim strTitle As String, strTmp As String
         For i = 2 To maxRow 'iLen + 1
            strTmp = .Cells(i, 1)
            strTitle = .Cells(i, 2)
            If valiRow(i - 2) > 0 Then
                dic.Add strTmp, strTitle
            End If
         Next i
    End With
      
    With Workbooks(bookName).Worksheets(sheetName)
     
        Dim readMode As String, strCell
        tempString1 = ""
        For i = 2 To 200
            strCell = .Cells(1, i)
             If dic.item(strCell) <> "" Then
                tempString2 = .Cells(1, i)
                tempString3 = dic.item(strCell)
                tempArray = 匹配数组列(tempString2)
                
                Select Case tempString3
                    Case Is = "int"
                        readMode = "ReadInt32()"
                    Case Is = "uint"
                        readMode = "ReadUInt32()"
                    Case Is = "float"
                        readMode = "ReadSingle()"
                    'Case Is = "script"
                    '    readMode = "ReadChars(tf.ReadInt32())"
                    'Case Is = "string"
                    Case Else
                        readMode = "ReadString()"
                End Select
                
                If tempArray(0) <> nil Then
                    tempString1 = tempString1 & tempArray(0) & "[" & tempArray(2) & "] " & " = tf." & readMode & ";" & vbCrLf & vbTab & vbTab & vbTab
                Else
                    If tempString3 = "script" Then   ' 这个还要确认
                        'tempString1 = tempString1 & tempString2 & " = ToLua.instance.funcBuffer.Add(tf.ReadString());" & vbCrLf & vbTab & vbTab & vbTab
                        tempString1 = tempString1 & tempString2 & " = LuaFramework.Util.LoadString(tf.ReadString());" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "int_array" Then
                        'tempString1 = tempString1 & tempString2 & " = Array.ConvertAll<string, int>(tf.Get<string>(__" & tempString2 & ").Split(new char[] { ',' }), str => int.Parse(str));" & vbCrLf & vbTab & vbTab & vbTab
                        'tempString1 = tempString1 & "if (__" & tempString2 & " == " & """" & """" & ") " & tempString2 & " = new int[0];" & vbCrLf & vbTab & vbTab & vbTab
                        'tempString1 = tempString1 & "else " & tempString2 & " = Array.ConvertAll<string, int>(tf.Get<string>(__" & tempString2 & ").Split(new char[] { ',' }), str => int.Parse(str));" & vbCrLf & vbTab & vbTab & vbTab
                        tempString1 = tempString1 & "{string s = tf." & readMode & ";if (s == """") " & tempString2 & " = new int[0];else " & tempString2 & " = Array.ConvertAll<string, int>(s.Split(new char[] { ',' }), str => int.Parse(str));}" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "float_array" Then
                        tempString1 = tempString1 & "{string s = tf." & readMode & ";if (s == """") " & tempString2 & " = new float[0];else " & tempString2 & " = Array.ConvertAll<string, float>(s.Split(new char[] { ',' }), str => float.Parse(str));}" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "string_array" Then
                        tempString1 = tempString1 & "{string s = tf." & readMode & ";if (s == """") " & tempString2 & " = new string[0];else " & tempString2 & " = Array.ConvertAll<string, string>(s.Split(new char[] { ',' }), str => str);}" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "enum" Then
                    
                    Else
                        tempString1 = tempString1 & tempString2 & " = tf." & readMode & ";" & vbCrLf & vbTab & vbTab & vbTab
                    End If
                End If
            End If
        Next
        tempString = Replace(tempString, "ID = tf.ReadInt32();", tempString1)
        End With
                
        ' ===== 生成 enum：public enum 行无缩进，其余行加 Tab =====
        Dim enumDefs As String
        enumDefs = ""
        
        If keyNum = 1 Then
            Dim key1FieldName As String: key1FieldName = ""
            Dim k As Long
            For k = 2 To maxRow
                If valiRow(k - 2) = 2 Then
                    key1FieldName = Workbooks("Settings.xls").Worksheets(sheetName & "Title").Cells(k, 1).Value
                    Exit For
                End If
            Next k
        
            Dim dataWs As Worksheet
            Set dataWs = Workbooks(bookName).Worksheets(sheetName)
            Dim dataMaxRow As Long: dataMaxRow = dataWs.UsedRange.rows.count
            Dim dataMaxCol As Long: dataMaxCol = dataWs.UsedRange.Columns.count
        
            ' ===【关键优化】整块读入内存，避免逐个 .Cells 调用 ===
            Dim dataArr As Variant
            If dataMaxRow >= 1 And dataMaxCol >= 1 Then
                dataArr = dataWs.Range(dataWs.Cells(1, 1), dataWs.Cells(dataMaxRow, dataMaxCol)).Value2
            Else
                GoTo SkipEnumGeneration
            End If
        
            Dim key1DataCol As Long: key1DataCol = 1
            Dim c As Long
            For c = 1 To dataMaxCol
                If CStr(dataArr(1, c)) = key1FieldName Then
                    key1DataCol = c
                    Exit For
                End If
            Next c
        
            For k = 2 To maxRow
                If valiRow(k - 2) = 1 Then
                    Dim fieldName As String, fieldType As String
                    With Workbooks("Settings.xls").Worksheets(sheetName & "Title")
                        fieldName = Trim(.Cells(k, 1).Value)
                        fieldType = Trim(.Cells(k, 2).Value)
                    End With
        
                    If StrComp(fieldType, "enum", vbTextCompare) = 0 And fieldName <> "" Then
                        Dim items As Object
                        Set items = CreateObject("System.Collections.ArrayList")
                        
                        Dim fieldDataCol As Long: fieldDataCol = 0
                        For c = 1 To dataMaxCol
                            If CStr(dataArr(1, c)) = fieldName Then
                                fieldDataCol = c
                                Exit For
                            End If
                        Next c
        
                        If fieldDataCol = 0 Then
                            enumDefs = enumDefs & "public enum " & fieldName & vbCrLf
                            enumDefs = enumDefs & vbTab & "{" & vbCrLf
                            enumDefs = enumDefs & vbTab & "    // ERROR: Column '" & fieldName & "' not found." & vbCrLf
                            enumDefs = enumDefs & vbTab & "}" & vbCrLf & vbCrLf
                        Else
                            Dim r As Long
                            For r = 2 To dataMaxRow
                                If CStr(dataArr(r, 1)) <> "*" Then
                                    GoTo NextRow
                                End If
        
                                Dim rawLabel As Variant: rawLabel = dataArr(r, fieldDataCol)
                                Dim rawKey As Variant: rawKey = dataArr(r, key1DataCol)
        
                                Dim labelStr As String: labelStr = ""
                                Dim keyStr As String: keyStr = ""
                                If Not IsEmpty(rawLabel) Then labelStr = Trim(CStr(rawLabel))
                                If Not IsEmpty(rawKey) Then keyStr = Trim(CStr(rawKey))
        
                                If labelStr <> "" And keyStr <> "" Then
                                    Dim cleanName As String: cleanName = ""
                                    Dim j As Long
                                    For j = 1 To Len(labelStr)
                                        Dim ch As String: ch = Mid(labelStr, j, 1)
                                        If ch Like "[A-Za-z_]" Or (j > 1 And ch Like "[0-9]") Then
                                            cleanName = cleanName & ch
                                        ElseIf ch Like "[ !-~]" Then
                                            If cleanName <> "" And Right(cleanName, 1) <> "_" Then
                                                cleanName = cleanName & "_"
                                            End If
                                        End If
                                    Next j
        
                                    Do While Len(cleanName) > 1 And Right(cleanName, 1) = "_"
                                        cleanName = Left(cleanName, Len(cleanName) - 1)
                                    Loop
        
                                    If cleanName = "" Then
                                        cleanName = "Item_" & r
                                    ElseIf cleanName Like "[0-9]*" Then
                                        cleanName = "_" & cleanName
                                    End If
        
                                    items.Add Array(cleanName, keyStr)
                                End If
NextRow:
                            Next r
        
                            enumDefs = enumDefs & "public enum " & fieldName & vbCrLf
                            enumDefs = enumDefs & vbTab & "{" & vbCrLf
        
                            If items.count = 0 Then
                                enumDefs = enumDefs & vbTab & "    // No valid items to export." & vbCrLf
                            Else
                                Dim maxLen As Long: maxLen = 0
                                Dim item
                                For Each item In items
                                    If Len(item(0)) > maxLen Then maxLen = Len(item(0))
                                Next item
        
                                Dim idx As Long
                                For idx = 0 To items.count - 1
                                    Dim Name As String: Name = items(idx)(0)
                                    Dim Value As String: Value = items(idx)(1)
                                    Dim padding As String: padding = Space(maxLen - Len(Name))
                                    Dim line As String
                                    line = "    " & Name & padding & " = " & Value
                                    If idx < items.count - 1 Then
                                        line = line & ","
                                    End If
                                    enumDefs = enumDefs & vbTab & line & vbCrLf
                                Next idx
                            End If
        
                            enumDefs = enumDefs & vbTab & "}" & vbCrLf & vbCrLf
                        End If
                    End If
                End If
            Next k
        End If
        
SkipEnumGeneration:
        
        If enumDefs = "" Then
            Dim regEx As Object
            Set regEx = CreateObject("VBScript.RegExp")
            With regEx
                .Global = True
                .Multiline = True
                .Pattern = "^[ \t]*//\[\[ENUM_DEFINITIONS\]\][ \t]*\r?\n?"
            End With
            tempString = regEx.Replace(tempString, "")
        Else
            tempString = Replace(tempString, "//[[ENUM_DEFINITIONS]]", enumDefs)
        End If
        
        Dim path1
        path1 = client & "\Assets\GameScript\ConfigInfo\Config\" & sheetName & ".cs"
        'path1 = "F:\tmp\" & sheetName & ".cs"
        'Call utf8SaveAs(path1, tempString)
        
        Call WriteUTF8File(tempString, path1, False)
        'Dim fso As FileSystemObject
        'Set fso = New FileSystemObject
        'Set file = fso.CreateTextFile(localPath & "\csConfig.cs", True)
        'file.Write tempString
        'file.Close
        
        'Dim table As Variant
        'table = .UsedRange.Value2
    
End Sub

'==============================
'导出.cpp文件
'==============================
Sub 导出C(sheetName As String, index As Long, server As String, ByVal pathForTxt As String)
    With Workbooks("Settings.xls").Worksheets(sheetName & "Title")
        '打开.CS配置文件模板
        Dim fso0 As FileSystemObject, tempFile, tempString As String, path As String
        'Set fso0 = New FileSystemObject
        'path = localPath & "\csConfigTemplate.txt"
        
        'tempFile = utf8Loadfile(path)
        'Set tempFile = fso0.OpenTextFile(path)
        
        Dim maxRow As Long
        maxRow = .UsedRange.rows.count
        
        Dim valiRow() As Long, valiRow2() As Long
        '哪几行标了*
        ReDim valiRow(maxRow - 1) As Long
        '哪几行有自定义导出
        ReDim valiRow2(maxRow - 1) As Long
        
        For i = 0 To maxRow - 2
            valiRow(i) = 0
        Next
        '找到首个*号所在列
        firstCol = 1
        Do
            firstCol = firstCol + 1
            temp = .Cells(2, firstCol)
        Loop Until temp = "*" Or temp = "key1" Or temp = "key2" Or temp = "key3"
        
        Dim keyNum As Long
        keyNum = 0
        For i = 2 To maxRow
            If .Cells(i, firstCol + 2 * (index - 1)) = "*" Then
                valiRow(i - 2) = 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key1" Then
                valiRow(i - 2) = 2
                keyNum = keyNum + 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key2" Then
                valiRow(i - 2) = 3
                keyNum = keyNum + 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key3" Then
                valiRow(i - 2) = 4
                keyNum = keyNum + 1
            End If
            If .Cells(i, firstCol + 1 + 2 * (index - 1)) <> nil Then
                valiRow2(i - 2) = 1
            End If
        Next i
        
        If keyNum = 1 Then
            path = localPath & "\KConfigTemplate1Manager_cpp.txt"
        ElseIf keyNum = 2 Then
            path = localPath & "\KConfigTemplate2Manager_cpp.txt"
        ElseIf keyNum = 3 Then
            path = localPath & "\KConfigTemplate3Manager_cpp.txt"
        End If
        tempFile = utf8Loadfile(path)
        
        'ReDim valiRow(maxRow - 1) As Long
        
        Dim tempString1 As String, tempString2, tempString3 As String, tempArray As Variant
        tempString = tempFile
        
        tempString = Replace(tempString, "[[tabname]]", sheetName)
        
        Dim keyFlag As Long
        keyFlag = 0
        For i = 2 To maxRow
            If valiRow(i - 2) = 2 Then
                tempString = Replace(tempString, "[[key1type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key1name]]", .Cells(i, 1))
                keyFlag = 1
            ElseIf valiRow(i - 2) = 3 Then
                tempString = Replace(tempString, "[[key2type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key2name]]", .Cells(i, 1))
            ElseIf valiRow(i - 2) = 4 Then
                tempString = Replace(tempString, "[[key3type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key3name]]", .Cells(i, 1))
            End If
        Next
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                tempArray = 匹配数组列(tempString2)
                'If tempArray(0) <> nil Then
                '    If tempArray(2) = 0 Then
                '        tempString1 = tempString1 & "public " & tempString3 & "[] " & tempArray(0) & " { get; private set; }" & vbCrLf & vbTab & vbTab
                '    End If
                'Else
                If tempArray(0) = nil Then
                    If tempString3 = "int" Or tempString3 = "uint" Or tempString3 = "script" Then
                        tempString1 = tempString1 & tempString2 & " = 0;" & vbCrLf & vbTab & vbTab
                'ElseIf tempString3 = "int_array" Then
                '    tempString1 = tempString1 & "public " & "int[]" & " " & tempString2 & " { get; private set; }" & vbCrLf & vbTab & vbTab
                    ElseIf tempString3 = "float" Then
                        tempString1 = tempString1 & tempString2 & " = 0.f;" & vbCrLf & vbTab & vbTab
                    End If
                'Else
                '    tempString1 = tempString1 & tempString2 & " = 0;" & vbCrLf & vbTab & vbTab
                End If
                'End If
            End If
        Next
        tempString = Replace(tempString, "[[colName]] = 0;", tempString1)
        
        'Debug.Print (tempString)
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) = 1 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                tempArray = 匹配数组列(tempString2)
                'If tempArray(0) <> nil Then
                If tempArray(0) <> nil Then
                    If tempString3 = "int" Then
                        tempString1 = tempString1 & "tabFile->GetInteger(""" & tempString2 & """, 0, &lnValue);" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "float" Then
                        tempString1 = tempString1 & "tabFile->GetFloat(""" & tempString2 & """, 0.f, &fValue);" & vbCrLf & vbTab & vbTab & vbTab
                    End If
                    tempString1 = tempString1 & "lpInfo->" & tempArray(0) & ".push_back(lnValue);" & vbCrLf & vbTab & vbTab & vbTab
                ElseIf tempString3 = "string" Then
                    tempString1 = tempString1 & "tabFile->GetString(""" & tempString2 & """, """", lszValue, _ConfigStringMax);" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "lpInfo->" & tempString2 & " = lszValue;" & vbCrLf & vbTab & vbTab & vbTab
                ElseIf tempString3 = "float" Then
                    tempString1 = tempString1 & "tabFile->GetFloat(""" & tempString2 & """, 0.f, &fValue);" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "lpInfo->" & tempString2 & " = fValue;" & vbCrLf & vbTab & vbTab & vbTab
                ElseIf tempString3 = "script" Then
                    tempString1 = tempString1 & "tabFile->GetString(""" & tempString2 & """, """", lszValue, _ConfigStringMax);" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "lpInfo->" & tempString2 & " = _LuaFunc(lszValue);" & vbCrLf & vbTab & vbTab & vbTab
                ElseIf tempString3 = "int_array" Or tempString3 = "float_array" Or tempString3 = "string_array" Then
                    tempString1 = tempString1 & "tabFile->GetString(""" & tempString2 & """, """", lszValue, _ConfigStringMax);" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "_SplitLine(lpInfo->" & tempString2 & ", lszValue);" & vbCrLf & vbTab & vbTab & vbTab
                ElseIf tempString3 = "int" And tempArray(0) = nil Then
                    tempString1 = tempString1 & "tabFile->GetInteger(""" & tempString2 & """, 0, &lnValue);" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "lpInfo->" & tempString2 & " = lnValue;" & vbCrLf & vbTab & vbTab & vbTab
                ElseIf tempString3 = "uint" And tempArray(0) = nil Then
                    tempString1 = tempString1 & "tabFile->GetInteger(""" & tempString2 & """, 0, &lnValue);" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "lpInfo->" & tempString2 & " = lnValue;" & vbCrLf & vbTab & vbTab & vbTab
                End If
            End If
        Next
        tempString = Replace(tempString, "tabFile->GetInteger(""[[colName]]"", 0, &lnValue);", tempString1)

        
        Dim path2
        path2 = server & "Src\CommonLogic\Config\"
        If Dir(path2, vbDirectory) <> "" Then
            Dim path1
            path1 = server & "Src\CommonLogic\Config\" & "K" & sheetName & ".cpp"
            'Call utf8SaveAs(path1, tempString)
            
            Call WriteUTF8File(tempString, path1, False)
            'Dim fso As FileSystemObject
            'Set fso = New FileSystemObject
            'Set file = fso.CreateTextFile(Workbooks("工具.xlam").path & "\csConfig.cs", True)
            'file.Write tempString
            'file.Close
            
            'Dim table As Variant
            'table = .UsedRange.Value2
        End If
               
    End With
End Sub

'==============================
'导出.h文件
'==============================
Sub 导出h(sheetName As String, index As Long, server As String, ByVal pathForTxt As String)
    With Workbooks("Settings.xls").Worksheets(sheetName & "Title")
        '打开.CS配置文件模板
        Dim fso0 As FileSystemObject, tempFile, tempString As String, path As String
        'Set fso0 = New FileSystemObject
        'path = localPath & "\csConfigTemplate.txt"
        
        'tempFile = utf8Loadfile(path)
        'Set tempFile = fso0.OpenTextFile(path)
        
        Dim maxRow As Long
        maxRow = .UsedRange.rows.count
        
        Dim valiRow() As Long, valiRow2() As Long
        '哪几行标了*
        ReDim valiRow(maxRow - 1) As Long
        '哪几行有自定义导出
        ReDim valiRow2(maxRow - 1) As Long
        
        For i = 0 To maxRow - 2
            valiRow(i) = 0
        Next
        '找到首个*号所在列
        firstCol = 1
        Do
            firstCol = firstCol + 1
            temp = .Cells(2, firstCol)
        Loop Until temp = "*" Or temp = "key1" Or temp = "key2" Or temp = "key3"
        
        Dim keyNum As Long
        keyNum = 0
        For i = 2 To maxRow
            If .Cells(i, firstCol + 2 * (index - 1)) = "*" Then
                valiRow(i - 2) = 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key1" Then
                valiRow(i - 2) = 2
                keyNum = keyNum + 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key2" Then
                valiRow(i - 2) = 3
                keyNum = keyNum + 1
            ElseIf .Cells(i, firstCol + 2 * (index - 1)) = "key3" Then
                valiRow(i - 2) = 4
                keyNum = keyNum + 1
            End If
            If .Cells(i, firstCol + 1 + 2 * (index - 1)) <> nil Then
                valiRow2(i - 2) = 1
            End If
        Next i
        
        If keyNum = 1 Then
            path = localPath & "\KConfigTemplate1Manager_h.txt"
        ElseIf keyNum = 2 Then
            path = localPath & "\KConfigTemplate2Manager_h.txt"
        ElseIf keyNum = 3 Then
            path = localPath & "\KConfigTemplate3Manager_h.txt"
        End If
        tempFile = utf8Loadfile(path)
        
        'ReDim valiRow(maxRow - 1) As Long
        
        Dim tempString1 As String, tempString2, tempString3 As String, tempArray As Variant
        tempString = tempFile
        
        tempString = Replace(tempString, "[[tabname]]", sheetName)
        
'        Dim reg As New RegExp
'        With reg
'            .Global = True
'            .IgnoreCase = True
'            '.Pattern = "\d+"
'            .Pattern = "\\Config\\.*$"
'        End With
'        Dim mc As MatchCollection
'        Set mc = reg.Execute(pathForTxt)
'        pathForTxt = mc.item(0)
'        pathForTxt = Replace(pathForTxt, "\", "/")
'
'        tempString = Replace(tempString, "[[path]]", pathForTxt)
        
        Dim keyFlag As Long
        keyFlag = 0
        For i = 2 To maxRow
            If valiRow(i - 2) = 2 Then
                tempString = Replace(tempString, "[[key1type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key1name]]", .Cells(i, 1))
                keyFlag = 1
            ElseIf valiRow(i - 2) = 3 Then
                tempString = Replace(tempString, "[[key2type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key2name]]", .Cells(i, 1))
            ElseIf valiRow(i - 2) = 4 Then
                tempString = Replace(tempString, "[[key3type]]", .Cells(i, 2))
                tempString = Replace(tempString, "[[key3name]]", .Cells(i, 1))
            End If
        Next
        
        
'        tempString1 = ""
'        For i = 2 To maxRow
'            If valiRow(i - 2) > 0 Then
'                tempString2 = .Cells(i, 1)
'                tempString1 = tempString1 & "public static readonly string __" & tempString2 & " = " & Chr(34) & tempString2 & Chr(34) & ";" & vbCrLf & vbTab & vbTab
'            End If
'        Next
'        tempString = Replace(tempString, "//public static readonly string __[[列名]] = " & """" & "[[列名]]" & """" & ";", tempString1)
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                tempArray = 匹配数组列(tempString2)
                If tempArray(0) <> nil Then
                    If tempArray(2) = 0 Then
                        tempString1 = tempString1 & "VECTOR_INT " & tempArray(0) & ";" & vbCrLf & vbTab & vbTab
                    End If
                ElseIf tempString3 = "string" Then
                    tempString1 = tempString1 & "KConfString " & tempString2 & ";" & vbCrLf & vbTab & vbTab
                ElseIf tempString3 = "int_array" Then
                    tempString1 = tempString1 & "VECTOR_INT " & tempString2 & ";" & vbCrLf & vbTab & vbTab
                ElseIf tempString3 = "float_array" Then
                    tempString1 = tempString1 & "VECTOR_FLOAT " & tempString2 & ";" & vbCrLf & vbTab & vbTab
                ElseIf tempString3 = "string_array" Then
                    tempString1 = tempString1 & "VECTOR_STRING " & tempString2 & ";" & vbCrLf & vbTab & vbTab
                ElseIf tempString3 = "script" Then
                    tempString1 = tempString1 & "int " & tempString2 & ";" & vbCrLf & vbTab & vbTab
                ElseIf tempString3 = "uint" Then
                    tempString1 = tempString1 & "UINT " & tempString2 & ";" & vbCrLf & vbTab & vbTab
                Else
                    tempString1 = tempString1 & tempString3 & " " & tempString2 & ";" & vbCrLf & vbTab & vbTab
                End If
                'End If
            End If
        Next
        tempString = Replace(tempString, "[[colType]] [[colName]]", tempString1)
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                If tempString3 = "string" Then
                    tempString1 = tempString1 & "const char* Get" & tempString2 & "() const" & vbCrLf & vbTab & vbTab
                    tempString1 = tempString1 & "{" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "return " & tempString2 & ".c_str();" & vbCrLf & vbTab & vbTab
                    tempString1 = tempString1 & "}" & vbCrLf & vbTab & vbTab
                End If
            End If
        Next
        tempString = Replace(tempString, "//const char* GetString()", tempString1)
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                If tempString3 = "int_array" Or tempString3 = "float_array" Or tempString3 = "string_array" Then
                    tempString1 = tempString1 & "int Get" & tempString2 & "Count() const" & vbCrLf & vbTab & vbTab
                    tempString1 = tempString1 & "{" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "return " & tempString2 & ".size();" & vbCrLf & vbTab & vbTab
                    tempString1 = tempString1 & "}" & vbCrLf & vbTab & vbTab
                End If
            End If
        Next
        tempString = Replace(tempString, "//const int GetIntArrayCount()", tempString1)
        
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) > 0 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                If tempString3 = "int_array" Then
                    tempString1 = tempString1 & "int Get" & tempString2 & "Value(int index) const" & vbCrLf & vbTab & vbTab
                    tempString1 = tempString1 & "{" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "return " & tempString2 & ".at(index);" & vbCrLf & vbTab & vbTab
                    tempString1 = tempString1 & "}" & vbCrLf & vbTab & vbTab
                ElseIf tempString3 = "float_array" Then
                    tempString1 = tempString1 & "float Get" & tempString2 & "Value(int index) const" & vbCrLf & vbTab & vbTab
                    tempString1 = tempString1 & "{" & vbCrLf & vbTab & vbTab & vbTab
                    tempString1 = tempString1 & "return " & tempString2 & ".at(index);" & vbCrLf & vbTab & vbTab
                    tempString1 = tempString1 & "}" & vbCrLf & vbTab & vbTab
                End If
            End If
        Next
        tempString = Replace(tempString, "//const int GetIntArrayValue()", tempString1)
        
        'Debug.Print (tempString)
        tempString1 = ""
        For i = 2 To maxRow
            If valiRow(i - 2) = 1 Then
                tempString2 = .Cells(i, 1)
                tempString3 = .Cells(i, 2)
                tempArray = 匹配数组列(tempString2)
                If tempArray(0) = nil Then
                    If tempString3 = "string" Then
                        tempString1 = tempString1 & "DefMemberFunc(Get" & tempString2 & ");" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "int" Or tempString3 = "uint" Or tempString3 = "float" Then
                        tempString1 = tempString1 & "DefMemberVar(" & tempString2 & ");" & vbCrLf & vbTab & vbTab & vbTab
                    ElseIf tempString3 = "int_array" Or tempString3 = "float_array" Then
                        tempString1 = tempString1 & "DefMemberFunc(Get" & tempString2 & "Count);" & vbCrLf & vbTab & vbTab & vbTab
                        tempString1 = tempString1 & "DefMemberFunc(Get" & tempString2 & "Value);" & vbCrLf & vbTab & vbTab & vbTab
                    End If
                End If
            End If
        Next
        tempString = Replace(tempString, "//DefMemberFunc(Get[[stringname]]);", tempString1)
        
        
        Dim path2
        path2 = server & "Src\CommonLogic\Config\"
        If Dir(path2, vbDirectory) <> "" Then
            Dim path1
            path1 = server & "Src\CommonLogic\Config\" & "K" & sheetName & ".h"
            'Call utf8SaveAs(path1, tempString)
            
            Call WriteUTF8File(tempString, path1, False)
            'Dim fso As FileSystemObject
            'Set fso = New FileSystemObject
            'Set file = fso.CreateTextFile(Workbooks("工具.xlam").path & "\csConfig.cs", True)
            'file.Write tempString
            'file.Close
            
            'Dim table As Variant
            'table = .UsedRange.Value2
        End If
               
    End With
End Sub

Function 匹配数组列(ByVal inputCol As String) As Variant
    Dim reg As New RegExp
    With reg
        .Global = True
        .IgnoreCase = True
        '.Pattern = "\d+"
        .Pattern = "^(.+)_([0-9]+)_([0-9]+)$"
    End With
    Dim mc As MatchCollection, temp(3) As Variant
    Set mc = reg.Execute(inputCol)

    For Each m In mc
        temp(0) = m.SubMatches(0)
        temp(1) = m.SubMatches(1)
        temp(2) = m.SubMatches(2)
    Next
    匹配数组列 = temp
End Function

'读取Settings获得列导出配置，存到一个Dictionary里，映射col_name->[datatype, msgid]
''data_type枚举：1=int/uint, 2=float, 3=string
'其中messageid = -1 为非message列，否则为message列的message段
Function GetColumnDef(ByVal sheetName As String)
    Dim maxRow, maxColumn, def(1) As Long            '[data_type,msgid]'
    Set GetColumnDef = CreateObject("Scripting.Dictionary")
    With Workbooks("Settings.xls").Sheets(sheetName & "Title")
        maxRow = .UsedRange.Columns.count
        maxColumn = .UsedRange.rows.count

        ' 生成dic（表头名，表头类型）
        
        Dim row As Long
        Dim p As Long
        Dim col_name, data_type, note As String

        p = 0
        ' 表头不再要了
        For row = 2 To maxColumn 'iLen + 1
            ' * || key 导出
            '构造字符串  一个字节的字符串长度+字符串
            col_name = .Cells(row, 1)

            ' 表头类型 只做表头类型，导出判断下面做
            data_type = .Cells(row, 2)

            ' 利用说明列，查询“MSG:”字样，以获得MessageID偏移量
            note = .Cells(row, 3)
            p = InStr(1, note, "MSG:")
            def(1) = -1                 '非message，def(1)默认为-1
            If p > 0 Then               '是message，def(1)为起始ID
                subStr = Right(note, Len(note) - (p + 4) + 1)
                def(1) = val(subStr)
            End If
                
            If col_name <> "" And data_type <> "" Then
                Select Case data_type
                    Case Is = "int"
                        def(0) = 1
                    Case Is = "uint"
                        def(0) = 1
                    Case Is = "float"
                        def(0) = 2
                    Case Is = "string"
                        def(0) = 3
                    Case Is = "enum"
                        def(0) = 4
                    Case Else               '其它暂时用string处理
                        def(0) = 3
                End Select
                Call GetColumnDef.Add(col_name, def)
            End If '   <>
        Next row
    End With
End Function

Sub DataCheck(ByVal 工作表 As String, ByVal 工作薄 As String)
    If 工作薄 = "" Then
        工作薄 = ActiveWorkbook.Name
    End If
    If 工作表 = "" Then
        工作薄 = ActiveWorkbook.Name
        工作表 = ActiveSheet.Name
    End If

    With Workbooks(工作薄).Sheets(工作表)
        Dim table As Variant
        table = .UsedRange.Value2
        
        Dim maxRow As Long
        maxRow = .UsedRange.rows.count
        
        Dim idCheckDict As New Dictionary
        Call idCheckDict.RemoveAll
 
        Dim i As Long, ID As String, nID As Long
        For i = 2 To maxRow
            If table(i, 1) = "*" Then
                ID = .Cells(i, 2)
                If nID > val(ID) Then
                    MsgBox 工作表 & " - ID 请按从小到大排列：" & ID, , "错误"
                    GoTo continue_i
                End If
                
                nID = val(ID)
                If idCheckDict.Exists(ID) Then
                    MsgBox 工作表 & " - ID 重复：" & ID, , "错误"
                    GoTo continue_i
                Else
                    Call idCheckDict.Add(ID, 1)
                End If
            End If
continue_i:
        Next
    End With
End Sub


' 导出xls 二进制数据文件，不带表头信息
' 注意，导出类型目前包含 int, uint, float, string, script，其他类型默认按string处理
' ！！！类型为string时，改长度支持，否则 这边二进制编码并没有做特殊处理，cs 文件将无法正确解析，不过script可以解析长度 2^31 -1
Public Sub new导出二进制文件(path As String, validCols() As Long, sheetName As String, bookName As String)
    ' 行数，列数
    Dim maxRow As Long, maxColumn As Long, firstColumn As Long
    
    'Dim valiRow() As Long
    '注意作用域  导出路径
     
    '需要保存单元格 信息表 下面存储单元格信息的时候需要知道每个单元格的类型（bookName，类型）
      
    On Error GoTo errHandle

    '从Settings.xlsx里获得列导出信息
    Dim colDefDict
    Set colDefDict = GetColumnDef(sheetName)

    ' ----------------------------------------------------------------
    ' 开始读取数据表，导出结果到目标文件path
    '构造有效列表，通过此表，判断后面的单元格是否需要导出
    'validRow = 读取导出列(1, sheetName, bookName)
    Dim ival As Long, fval As Single, strTmp As String, cell_id As String, cell_value As String, bTmp As Byte
    
    maxColumn = UBound(validCols) - LBound(validCols) + 1
    With Workbooks(bookName).Sheets(sheetName)
        '如果目标文件存在，删除目标文件
        If Dir(path, vbDirectory) <> Empty Then
            Kill path
        End If
        
        '打开目标文件作为二进制写入
        Open path For Binary Access Write As #1
        maxRow = .UsedRange.rows.count
        firstColumn = 0

        For row = 2 To maxRow
            ' 判断该行是否需要导出 Skill.xls 每一行第一列都是 以 * 打头判断
            If .Cells(row, 1) <> "*" Then
                 GoTo continue_row
            End If
            ' 做了加2处理，上界变为iLen-1 -> iLen+1
            For col = 2 To maxColumn
                ' 判断该列是否需要导出
                If validCols(col - 2) <> 1 Then
                    GoTo continue_col
                End If

                If firstColumn = 0 Then
                    firstColumn = col
                End If
                
                colName = .Cells(1, col)
                Dim def
                def = colDefDict.item(colName)
                data_type = def(0)
                msgId = def(1)
                
                cell_id = .Cells(row, firstColumn)
                cell_value = .Cells(row, col)
                'msgid <> -1表示该列是message列，导出message_id并更新message
                If msgId <> -1 Then
                    ival = msgId + cell_id
                    Call AddMessage(ival, cell_value)
                    Put #1, , ival
                Else
                    Select Case data_type
                        Case Is = 1     ' int,uint
                            ' 如果空单元格里面隐藏了有空单元格，得处理一下
                            ' 用 val(str) 能处理
                            If IsEmpty(cell_value) Then
                                ival = 0
                            Else
                                iTmp = Len(cell_value)
                                If iTmp <= 9 Then
                                    ival = val(cell_value)
                                Else
                                    iival = val(cell_value)
                                    ' 长整型数据类型的取值范围是 -2^31 ~ 2^31-1 也就是(-2147483648 ~ 2147483647)
                                    If iival > 2147483647 Then
                                        ival = -(2147483647 - (iival - 2147483647)) - 2
                                    Else
                                        ival = val(cell_value)
                                    End If
                                End If
                            End If
                            Put #1, , ival
                        Case Is = 2     ' float
                            If IsEmpty(cell_value) Then
                                fval = 0
                            Else
                                fval = val(cell_value)
                            End If
                            Put #1, , fval
                        Case Is = 4     ' enum

    
                        Case Else ' 其它类型都默认用string 处理，程度再负责数据转换 现在又出现中文，也是
                            If IsEmpty(cell_value) Or cell_value = "" Then
                                bTmp = 0
                                strTmp = ""
                                Put #1, , bTmp
                                Put #1, , strTmp
                            Else
                                'bb = StringToUTF8Binary(cell_value)
                                'Put #1, , bb
                                iTmp = Len(cell_value)
                                ' 转成utf8编码
                                Dim bb() As Byte
                                bb = Utf8BytesFromString(cell_value)
                                iTmp = UBound(bb) - LBound(bb) + 1
                                
                                Do While iTmp >= 128
                                    bTmp = iTmp Mod 128 + 128
                                    Put #1, , bTmp
                                    iTmp = Int(iTmp / 128)
                                Loop
                                bTmp = iTmp
                                Put #1, , bTmp
                                Put #1, , bb

                            End If
                    End Select
                End If
continue_col:
            Next col
continue_row:
        Next row
        Close #1
    End With
    
    Exit Sub
errHandle:
    MsgBox Err.Description & " sheetName(" & sheetName & "), ID(" & cell_id & "), 列名(" & colName & ")", , "错误 - " & Err.Number
End Sub


Sub new导出TXT(ByVal 路径 As String, valiRow() As Long, ByVal 星号 As Long, ByVal 工作表 As String, ByVal 工作薄 As String)

    'Call frmLog.Log("正在导出 " & 工作薄 & "下" & 工作表 & "到 " & 路径)
    
    If 工作薄 = "" Then
        工作薄 = ActiveWorkbook.Name
    End If
    If 工作表 = "" Then
        工作薄 = ActiveWorkbook.Name
        工作表 = ActiveSheet.Name
    End If

    '从Settings.xlsx里获得列导出信息
    Dim def
    Dim colDefDict
    Set colDefDict = GetColumnDef(工作表)


    Dim tempString As String
    With Workbooks(工作薄).Sheets(工作表)
        Dim table As Variant
        table = .UsedRange.Value2
        Dim i As Long, j As Long
        
        Dim maxRow As Long
        Dim maxCol As Long
        maxRow = .UsedRange.rows.count
        maxCol = .UsedRange.Columns.count
        Dim valiColumn() As Long
        
        valiColumn = valiRow
        
        '新导出方法
        
       
        Dim follow As String, firstCol As Long, ID As String, msgId As Long

        For i = 1 To maxRow
            firstCol = 0
            'If Not IsEmpty(table(i, 星号)) Then
            If table(i, 1) = "*" Then
                For j = 2 To maxCol
                    colName = .Cells(1, j)
                    def = colDefDict.item(colName)
                    If (valiColumn(j - 2) = 1) Then
                        If firstCol = 0 Then
                            firstCol = j
                        Else
                            tempString = tempString & vbTab
                        End If
                        
                        ID = .Cells(i, firstCol)
                        If i <> 1 And def(1) <> -1 Then           '这是message string，转成MessageID
                            msgId = def(1) + .Cells(i, firstCol)
                            tempString = tempString & CStr(msgId)
                        ElseIf table(i, j) = "" Then
                            tempString = tempString & ""
                        Else
                            tempString = tempString & table(i, j)
                        End If
                    End If
                Next
                If i <> maxRow Then
                    tempString = tempString & vbCrLf
                End If
            End If
            'End If
continue_i:
        Next
        
    End With
    'file.Close
    'Set file = Nothing
    'Set fso = Nothing
    'Call utf8SaveAs(路径, tempString)
    Call WriteUTF8File(tempString, 路径, False)
    Call frmLog.Log("导出 " & 工作薄 & "下" & 工作表 & "到 " & 路径 & " 成功!")

End Sub


Sub 自由导出()
    Dim a As String
    a = BrowDir()
    If a = "" Then Exit Sub
    导出TXT a & ActiveSheet.Name & ".txt", 1
End Sub


'根据XML中Message_Info表的定义，得到客户端、服务器的Message文件路径，然后调用LoadMessage读取Message表
Sub InitMessage(ByRef dom As DOMDocument60, ByVal server As String, ByVal client As String, ByVal Distribute As String)
    msgClientPath = ""
    enMsgClientPath = ""
    msgServerPath = ""
    msgDistPath = ""
    For Each Node In dom.SelectNodes("Settings/Item[@Sheet='Language_cn_Info']/Export/File")
        If InStr(Node.text, "[Client]") = 1 Then
            msgClientPath = 替换(Node.text, "[Client]", client)
            msgClientPath = Replace(msgClientPath, ".txt", ".bin")
            enMsgClientPath = Replace(msgClientPath, "_cn", "_en")
        ElseIf InStr(Node.text, "[Server]") = 1 Then
            msgServerPath = 替换(Node.text, "[Server]", server)
        ElseIf InStr(Node.text, "[Distribute]") = 1 Then
            msgDistPath = 替换(Node.text, "[Distribute]", Distribute)
        End If
    Next Node
    
    Call messageDict.RemoveAll
    Call englishDict.RemoveAll
    MsgCount = 0
    Call msgBases.RemoveAll
End Sub

'读取Settings.xls内"MSG:"数量
Sub InitMessageCount(ByVal sheetName As String)
    Dim maxColumn, def(1) As Long
    With Workbooks("Settings.xls").Sheets(sheetName & "Title")
        maxColumn = .UsedRange.rows.count
        Dim row As Long
        Dim p As Long
        Dim note As String

        p = 0
        For row = 2 To maxColumn 'iLen + 1
            note = .Cells(row, 3)
            p = InStr(1, note, "MSG:")
            If p > 0 Then
                MsgCount = MsgCount + 1
                subStr = Right(note, Len(note) - (p + 4) + 1)
                Call msgBases.Add(val(subStr), 1)
            End If
        Next row
    End With
End Sub

'读取之前导出到服务器的Message.txt，构建messageDict
Sub LoadMessage()
    Dim strLine, strMsg As String, msgLines, msgPair() As String, ID As Long, idBase As Long
    strMsg = ""
    If msgServerPath <> "" And Dir(msgServerPath, vbDirectory) <> Empty Then
        strMsg = utf8Loadfile(msgServerPath)
    ElseIf msgDistPath <> "" And Dir(msgDistPath, vbDirectory) <> Empty Then
        strMsg = utf8Loadfile(msgDistPath)
    End If
    
   
    '处理每一行
    Dim firstRow As Boolean
    firstRow = True
    For Each strLine In Split(strMsg, vbCrLf)
        If Not firstRow Then
            msgPair = Split(strLine, vbTab)
            If UBound(msgPair) >= 1 Then
                ID = val(msgPair(0))
                '如果是本地message，则不加载
                If Not IsLocalMessage(ID) Then
                     Call messageDict.Add(ID, Replace(msgPair(1), ",", "，"))
                End If
                If UBound(msgPair) > 1 Then
                    Call englishDict.Add(ID, CStr(msgPair(2)))
                End If
            End If
        End If
        firstRow = False
    Next strLine
End Sub

'判断此messageID是否属于当前导出表单
'这支持msgIdBase为7个0或少于7个0的情况
Function IsLocalMessage(msgId As Long)
    For Each idBase In msgBases.keys
        Dim idRange As Long
        idRange = 10000000              '最多7个0
        While idBase Mod idRange > 0
            idRange = idRange / 10
        Wend
        If msgId >= idBase And msgId < idBase + idRange Then
            IsLocalMessage = True
            Exit Function
        End If
    Next idBase
    IsLocalMessage = False
End Function


'添加一个Message
Sub AddMessage(ID As Long, ByVal message As String)
    Dim idBase As Long
    If message <> "" Then
        '添加
        If messageDict.Exists(ID) Then
            MsgBox "ID 重复：" & " ID=" & ID, "错误"
        Else
            Call messageDict.Add(ID, Replace(message, ",", "，"))
        End If
    End If
End Sub

Sub QuickSort(ByRef arr() As Variant, ByVal low As Long, ByVal high As Long)
    If low < high Then
        Dim pivot As Variant
        Dim i As Long, j As Long
        Dim temp As Variant
        
        pivot = arr((low + high) \ 2)
        i = low
        j = high
        
        Do While i <= j
            Do While arr(i) < pivot
                i = i + 1
            Loop
            Do While arr(j) > pivot
                j = j - 1
            Loop
            If i <= j Then
                ' 交换键
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
                i = i + 1
                j = j - 1
            End If
        Loop
        
        QuickSort arr, low, j
        QuickSort arr, i, high
    End If
End Sub

'将messageDict的内容同时写入到服务器Message.txt和客户端Message.bin
Sub SaveMessage()
    Dim strTmp As String, path As String, i, j As Long, bTmp As Byte
        
    Dim keys() As Variant
    keys = messageDict.keys
    QuickSort keys, LBound(keys), UBound(keys)

    '先将txt存入Server/Distribute目标
    Dim paths(1) As String
    paths(0) = msgDistPath
    paths(1) = msgServerPath
    For i = 0 To 1
        If paths(i) <> "" Then
            'title
            If englishDict.count > 0 Then
                strTmp = "ID" & vbTab & "String" & vbTab & "english" & vbCrLf
            Else
                strTmp = "ID" & vbTab & "String" & vbCrLf
            End If

            For j = LBound(keys) To UBound(keys)
            'For j = 0 To messageDict.count - 1
                ID = keys(j)
                tempMsg = messageDict.item(ID)
                tempMsg = Replace(tempMsg, vbLf, "\r\n")

                If englishDict.count > 0 Then
                    enMsg = ""
                    If englishDict.Exists(ID) Then
                        enMsg = englishDict.item(ID)
                    End If
                    strTmp = strTmp & CStr(ID) & vbTab & tempMsg & vbTab & CStr(enMsg) & vbCrLf
                Else
                    strTmp = strTmp & CStr(ID) & vbTab & tempMsg & vbCrLf
                End If
            Next j
            Call WriteUTF8File(strTmp, paths(i), False)
        End If
    Next i

    '二进导出到客户端Message文件
    If msgClientPath <> "" Then
        If Dir(msgClientPath, vbDirectory) <> Empty Then
            Kill msgClientPath
        End If
        Open msgClientPath For Binary Access Write As #1
        For j = LBound(keys) To UBound(keys)
        'For j = 0 To messageDict.count - 1
            ID = keys(j)
            strTmp = messageDict.item(ID)
            strTmp = Replace(strTmp, vbLf, "\r\n")
            Dim TempLng As Long
            TempLng = ID
            Put #1, , TempLng
            
            'Put #1, , StringToUTF8Binary(strTmp)
            If IsEmpty(strTmp) Or strTmp = "" Then
                bTmp = 0
                strTmp = ""
                Put #1, , bTmp
                Put #1, , strTmp
            Else
                iTmp = Len(strTmp)
                ' 转成utf8编码
                Dim bb() As Byte
                bb = Utf8BytesFromString(strTmp)
                iTmp = UBound(bb) - LBound(bb) + 1
                
                Do While iTmp >= 128
                    bTmp = iTmp Mod 128 + 128
                    Put #1, , bTmp
                    iTmp = Int(iTmp / 128)
                Loop
                bTmp = iTmp
                Put #1, , bTmp
                Put #1, , bb
            End If
            
        Next j
        Close #1
    End If
End Sub

'将englishDict的内容同时写入到客户端en.bin
Sub SaveEnglishMessage()
    Dim strTmp As String, path As String, i, j As Long, bTmp As Byte

    Dim keys() As Variant
    keys = messageDict.keys
    QuickSort keys, LBound(keys), UBound(keys)

    '二进导出到客户端Message文件
    If enMsgClientPath <> "" Then
        If Dir(enMsgClientPath, vbDirectory) <> Empty Then
            Kill enMsgClientPath
        End If
        Open enMsgClientPath For Binary Access Write As #1
        For j = LBound(keys) To UBound(keys)
        'For j = 0 To messageDict.count - 1
            ID = keys(j)
            strTmp = ""
            If englishDict.Exists(ID) Then
                strTmp = englishDict.item(ID)
            End If
            'strTmp = messageDict.item(id)
            strTmp = Replace(strTmp, vbLf, "\r\n")
            Dim TempLng As Long
            TempLng = ID
            Put #1, , TempLng
            
            'Put #1, , StringToUTF8Binary(strTmp)
            If IsEmpty(strTmp) Or strTmp = "" Then
                bTmp = 0
                strTmp = ""
                Put #1, , bTmp
                Put #1, , strTmp
            Else
                iTmp = Len(strTmp)
                ' 转成utf8编码
                Dim bb() As Byte
                bb = Utf8BytesFromString(strTmp)
                iTmp = UBound(bb) - LBound(bb) + 1
                
                Do While iTmp >= 128
                    bTmp = iTmp Mod 128 + 128
                    Put #1, , bTmp
                    iTmp = Int(iTmp / 128)
                Loop
                bTmp = iTmp
                Put #1, , bTmp
                Put #1, , bb
            End If
            
        Next j
        Close #1
    End If
End Sub


'将字符串转换成C#可读取的UTF-8二进制流
Function StringToUTF8Binary(ByVal text As String)
    Dim bTmp As Byte, buff() As Byte, buffTxt
    Dim i As Long, j As Long, txtLen As Long
    
    If IsEmpty(text) Or text = "" Then
        ReDim buff(2)
        bTmp = 0
        buff(0) = bTmp
        bTmp = strTmp
        buff(1) = bTmp
    Else
        ' 先转成utf8编码
        buffTxt = Utf8BytesFromString(text)
        txtLen = UBound(buffTxt) - LBound(buffTxt) + 1
        
        ReDim buff(Int(txtLen / 128) + txtLen)
                        
        '写入长度信息
        i = 0
        Do While txtLen > 0
            If txtLen > 128 Then
                bTmp = txtLen Mod 128 + 128
            Else
                bTmp = txtLen
            End If
            buff(i) = bTmp
            txtLen = Int(txtLen / 128)
            i = i + 1
        Loop
        
        '再附上具体字符串
        txtLen = UBound(buffTxt) - LBound(buffTxt) + 1
        For j = LBound(buffTxt) To UBound(buffTxt)
            buff(i) = buffTxt(j)
            i = i + 1
        Next j
    End If
    
    StringToUTF8Binary = buff
    
End Function

#If Win64 Then
    Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
            ByVal CodePage As Long, _
            ByVal dwFlags As Long, _
            ByRef lpMultiByteStr As Any, _
            ByVal cchMultiByte As Long, _
            ByVal lpWideCharStr As Long, _
            ByVal cchWideChar As Long) As Long
    Public Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
            ByVal CodePage As Long, _
            ByVal dwFlags As Long, _
            ByVal lpWideCharStr As LongPtr, _
            ByVal cchWideChar As Long, _
            ByVal lpMultiByteStr As LongPtr, _
            ByVal cchMultiByte As Long, _
            ByVal lpDefaultChar As LongPtr, _
            ByVal lpUsedDefaultChar As LongPtr) As Long
        
#Else
    Public Declare Function MultiByteToWideChar Lib "kernel32" ( _
            ByVal CodePage As Long, _
            ByVal dwFlags As Long, _
            ByRef lpMultiByteStr As Any, _
            ByVal cchMultiByte As Long, _
            ByVal lpWideCharStr As Long, _
            ByVal cchWideChar As Long) As Long
    Public Declare Function WideCharToMultiByte Lib "kernel32" ( _
            ByVal CodePage As Long, _
            ByVal dwFlags As Long, _
            ByVal lpWideCharStr As Long, _
            ByVal cchWideChar As Long, _
            ByRef lpMultiByteStr As Any, _
            ByVal cchMultiByte As Long, _
            ByVal lpDefaultChar As String, _
            ByVal lpUsedDefaultChar As Long) As Long
#End If
        
Public Const CP_UTF8 = 65001
' 将输入文本写进UTF8格式的文本文件
' 输入
' strInput：文本字符串
' strFile：保存的UTF8格式文件路径
' bBOM：True表示文件带"EFBBBF"头，False表示不带
Sub WriteUTF8File(strInput As String, ByVal strFile As String, Optional bBOM As Boolean = True)
    Dim bByte As Byte
    Dim ReturnByte() As Byte
    Dim lngBufferSize As Long
    Dim lngResult As Long
    Dim TLen As Long
 
    ' 判断输入字符串是否为空
    If Len(strInput) = 0 Then Exit Sub
    On Error GoTo errHandle
    ' 判断文件是否存在，如存在则删除
    If Dir(strFile) <> "" Then Kill strFile
 
    TLen = Len(strInput)
    lngBufferSize = TLen * 3 + 1
    ReDim ReturnByte(lngBufferSize - 1)
    #If Win64 Then
        lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strInput), TLen, _
            VarPtr(ReturnByte(0)), lngBufferSize, 0&, 0&)
    #Else
        lngResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(strInput), TLen, _
            ReturnByte(0), lngBufferSize, vbNullString, 0)
    #End If
    If lngResult Then
        lngResult = lngResult - 1
        ReDim Preserve ReturnByte(lngResult)
        Open strFile For Binary As #1
        If bBOM = True Then
            bByte = 239
            Put #1, , bByte
            bByte = 187
            Put #1, , bByte
            bByte = 191
            Put #1, , bByte
        End If
        Put #1, , ReturnByte
        Close #1
    End If
    Exit Sub
errHandle:
    MsgBox Err.Description, , "错误 - " & Err.Number
End Sub



' basUtf8FromString

' Written by David Ireland DI Management Services Pty Limited 2015
' <http://www.di-mgt.com.au> <http://www.cryptosys.net>

Option Explicit

''' WinApi function that maps a UTF-16 (wide character) string to a new character string
#If Win64 Then
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long
#Else
    Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As Long, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long
#End If
' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001

''' Return byte array with VBA "Unicode" string encoded in UTF-8
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function
