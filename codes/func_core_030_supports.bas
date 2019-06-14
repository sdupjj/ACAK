Attribute VB_Name = "func_core_030_supports"
'1-----------------------------------------------
'1   Very basic function / sub in ACAK
'1   cs_ = core support
'1   cv_ = core variables
'1-----------------------------------------------
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Function cs_num2asc2(n As Long) As String
'1程序功能：
'1   输入数字返回英语字母
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_n As Long
'2赋值
    cv_n = n
'2代码
    cs_num2asc2 = Mid(Cells(1, cv_n).Address, 2, IIf(cv_n < 27, 1, 2))
End Function

Public Function cs_FV(Ipt As Variant, Optional WhichSheet As String = "core_setup", Optional WhichCol01 As String = "F", _
                Optional WhichCol02 As String = "G", Optional HowManyRows As Long = 1000) As Variant
'1程序功能：
'1   在指定的表中，通过指定列中的值，找另一个指定列中的值
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_i As Long
    Dim cv_x As Long
    Dim cv_n As Long
    Dim cv_InputValue As Variant
    Dim cv_SheetsName As String
    Dim cv_Col01Name As String
    Dim cv_Col02Name As String
    Dim cv_HowManyRowsInCol01 As Long
'2赋值
    cv_InputValue = Ipt
    cv_SheetsName = WhichSheet
    cv_Col01Name = WhichCol01
    cv_Col02Name = WhichCol02
    cv_HowManyRowsInCol01 = HowManyRows
'2代码
    With ThisWorkbook.Sheets(cv_SheetsName)
        For cv_x = 1 To .Range(cv_Col01Name & cv_HowManyRowsInCol01).End(xlUp).Row
            If .Range(cv_Col01Name & cv_x).value = cv_InputValue Then
                cs_FV = .Range(cv_Col02Name & cv_x).value
                Exit For
            End If
        Next
    End With
End Function
Public Sub cs_WV(Ipt As Variant, Wpt As Variant, Optional WhichSheet As String = "core_setup", Optional WhichCol01 As String = "F", _
                Optional WhichCol02 As String = "G", Optional HowManyRows As Long = 1000)
'1程序功能：
'1   在指定的表中，通过指定列中的值Ipt，写入另一个指定列中的值Wpt
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_i As Long
    Dim cv_x As Long
    Dim cv_n As Long
    Dim cv_InputValue As Variant
    Dim cv_WriteValue As Variant
    Dim cv_SheetsName As String
    Dim cv_Col01Name As String
    Dim cv_Col02Name As String
    Dim cv_HowManyRowsInCol01 As Long
'2赋值
    cv_InputValue = Ipt
    cv_WriteValue = Wpt
    cv_SheetsName = WhichSheet
    cv_Col01Name = WhichCol01
    cv_Col02Name = WhichCol02
    cv_HowManyRowsInCol01 = HowManyRows
'2代码
    With ThisWorkbook.Sheets(cv_SheetsName)
        For cv_x = 1 To .Range(cv_Col01Name & cv_HowManyRowsInCol01).End(xlUp).Row
            If .Range(cv_Col01Name & cv_x).value = cv_InputValue Then
                    .Range(cv_Col02Name & cv_x).value = cv_WriteValue
                Exit For
            End If
        Next
    End With
End Sub

Public Sub cs_AddOne(Whichcount As String)
'1程序功能：
'1   为count原有的值增加1
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_MiddleCount As Long
    Dim cv_WhichCount As String
'2赋值
    cv_WhichCount = Whichcount
'2代码
    cv_MiddleCount = cs_FV(cv_WhichCount, cs_FV("CountSheet"), "A", "B")
    Call cs_WV(cv_WhichCount, cv_MiddleCount + 1, cs_FV("CountSheet"), "A", "B")
End Sub

Public Sub cs_ReduceOne(Whichcount As String)
'1程序功能：
'1   为count原有的值减少1
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_MiddleCount As Long
    Dim cv_WhichCount As String
'2赋值
    cv_WhichCount = Whichcount
'2代码
    cv_MiddleCount = cs_FV(cv_WhichCount, cs_FV("CountSheet"), "A", "B")
    Call cs_WV(cv_WhichCount, cv_MiddleCount - 1, cs_FV("CountSheet"), "A", "B")
End Sub

Public Sub cs_BeZero(Whichcount As String)
'1程序功能：
'1   设置count的值为0
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_WhichCount As String
'2赋值
    cv_WhichCount = Whichcount
'2代码
    Call cs_WV(cv_WhichCount, 0, cs_FV("CountSheet"), "A", "B")
End Sub

Public Sub cs_BeOne(Whichcount As String)
'1程序功能：
'1   设置count的值为1
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_WhichCount As String
'2赋值
    cv_WhichCount = Whichcount
'2代码
    Call cs_WV(cv_WhichCount, 1, cs_FV("CountSheet"), "A", "B")
End Sub

Public Sub cs_BeValue(Whichcount As String, value As Variant)
'1程序功能：
'1   写入内容页里的 whichcount 旁边的数值变成value
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_WhichCount As String
    Dim cv_Value As Variant
'2赋值
    cv_WhichCount = Whichcount
    cv_Value = value
'2代码
    Call cs_WV(cv_WhichCount, cv_Value, cs_FV("CountSheet"), "A", "B")
End Sub

Public Function cs_FindValue(Whichcount As String) As Variant
'1程序功能：
'1   发现的whichcount 旁边的数值
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_WhichCount As String
'2赋值
    cv_WhichCount = Whichcount
'2代码
    cs_FindValue = cs_FV(cv_WhichCount, cs_FV("CountSheet"), "A", "B")
End Function

Public Sub cs_TakeAction(whichpagename As String, whichcol As String, Optional Rows As Long = 2, Optional RowE As Long = 999, _
                                            Optional ColorCellsOrNot As String = "N", Optional ShowErrorOrNot As String = "N")
'1程序功能：
'1   对某页的某列动作 进行执行
'1程序版本：
'1   1.1
'1版本修订：
'1   1.0 >>> 原始版本,并无颜色标记
'1   1.1 >>> 在状态栏显示动作
'2定义
    Dim cv_RowE As Long
    Dim cv_AA As String
    Dim cv_i As Long
    Dim cv_ActionsCamp As String  'actions in which vba module
    Dim cv_WhichPage As String
    Dim cv_WhichCol As String
    Dim cv_Rows As Long
    Dim cv_WorkbookName As String
    Dim cv_ColorCellsOrNot As String
    Dim cv_ShowErrorOrNot As String
    Dim cv_T01
    Dim cv_lindex As Long
    Dim cv_A As String
    Dim cv_B As String
    Dim cv_C As String
    Dim cv_NeedfoundString As String
    
'2赋值

    cv_WhichPage = whichpagename
    cv_WhichCol = whichcol
    cv_Rows = Rows
    cv_ShowErrorOrNot = ShowErrorOrNot
    cv_ColorCellsOrNot = ColorCellsOrNot
    cv_B = ""
    If RowE = 999 Then
        cv_RowE = ThisWorkbook.Sheets(cv_WhichPage).Range(cv_WhichCol & 100).End(xlUp).Row
    Else
        cv_RowE = RowE
    End If
    cv_ActionsCamp = cs_FV("ActionsInWhichVBAModule")
    cv_ColorCellsOrNot = ColorCellsOrNot
'2代码
'2得到可用的所有action
    With ThisWorkbook.VBProject.VBComponents(cs_FV("ActionsInWhichVBAModule")).CodeModule
        For cv_lindex = 1 To .CountOfLines
            cv_A = .Lines(cv_lindex, 1)
            cv_NeedfoundString = "Public Sub "
            If InStr(cv_A, cv_NeedfoundString) > 0 Then
                cv_A = Replace(cv_A, "Public Sub ", "")
                cv_B = cv_B & cv_A
            End If
        Next
    End With
    
    For cv_i = cv_Rows To cv_RowE
        With ThisWorkbook.Sheets(cv_WhichPage).Range(cv_WhichCol & cv_i)
            If cv_ColorCellsOrNot = "Y" Then
                .Interior.Color = RGB(207, 1, 37)
                .Font.Color = RGB(255, 255, 255)
            End If
            '2 得到当前要执行的action名字
            If InStr(.value, " ") > 0 Then
                cv_C = Left(.value, InStr(.value, " ") - 1)
            Else
                cv_C = .value
            End If
            '2查看下这个要执行的action是否在模块中可用被发现，如果不可以则在日志中写下error
            If InStr(cv_B, cv_C) = 0 Then
                Call cs_Log(.value & " can not be found!", "Error")
            End If
            cv_AA = "'" & cv_ActionsCamp & "." & .value & "'"
            cv_T01 = Timer
            If cv_ShowErrorOrNot = "Y" Then
                Call cs_Log(cv_AA, "Info")
                Application.Run cv_AA
            Else
                cs_runAA (cv_AA)
            End If
            Call cs_Log(.value & " 耗时： " & Str(Application.Text((Timer - cv_T01), "0.0000")), "Debug")
            DoEvents
            If cv_ColorCellsOrNot = "Y" Then
                .Interior.Color = RGB(255, 255, 255)
                .Font.Color = RGB(0, 0, 0)
            End If
        End With
    Next
    Exit Sub

End Sub

Public Sub cs_runAA(AA As String)
'1程序功能：
'1   运行一段代码
'1程序版本：
'1   1.1
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
Dim cv_AA As String
'2赋值
cv_AA = AA
'2代码
On Error GoTo debugprint
    Application.Run cv_AA
    Call cs_Log("Run" & cv_AA, "Info")
    Exit Sub
debugprint:
    Call cs_Log("Run" & cv_AA & " works error.", "Error")
End Sub

Public Sub cs_Log(sts As Variant, Optional Level_Error_Warning_Debug_Info_Print As String = "Print")
'1程序功能：
'1   记录sts到core_log中
'1程序版本：
'1   1.0
'1   1.1
'1   1.2
'1版本修订：
'1   1.0 >>> 原始版本
'1   1.1 >>> 增加一个参数 确定 是在什么情况下记录log
'1    1.2>>> 将原来需要分散设置 日志记录类型Level_Error_Warning_Debug_Info_Print 现在集中在这个程序中配置
'2定义
    Dim cv_PageName As String
    Dim cv_A As Long
    Dim cv_B As String
    Dim cv_previousText As String
    Dim cv_MaxLineInInfoText As Long
    Dim cv_i As Long
    Dim cv_ii As Long
    Dim cv_iii As Long
    Dim cv_n As Long
    Dim cv_nn As Long
    Dim cv_Record As Variant
    Dim cv_LogLevel As String
    Dim cv_array01() As String
    Dim cv_array02() As String
    Dim cv_array03() As String
'2赋值
    cv_PageName = cs_FV("LogSheet")
    cv_Record = sts
    cv_LogLevel = Level_Error_Warning_Debug_Info_Print
    cv_array01() = Split(cs_FV("LogLevel"), ";")
    cv_i = UBound(cv_array01)
    cv_array02() = Split(cs_FV("ShowInInfoWindowsLogLevel"), ";")
    cv_ii = UBound(cv_array02)
'2代码
'2在core_log 页里面显示
    For cv_n = 0 To cv_i
        If cv_array01(cv_n) = cv_LogLevel Then
            '2记录log记录到第几行
            '2 count03  这个值应该在ACAK启动或关闭时被重新设置成 0 ！！！
            Call cs_AddOne("count03")
            cv_A = cs_FindValue("count03") + 1
            '2生成log记录
            cv_B = Now() & " <" & cv_LogLevel & "> " & cv_Record
            '2保存到log记录页中
            ThisWorkbook.Sheets(cv_PageName).Range("B" & cv_A).value = cv_B
            Exit For
        End If
    Next
'2在info窗体里面显示
    If Info.Visible = True Then
        For cv_nn = 0 To cv_ii
            If cv_array02(cv_nn) = cv_LogLevel Then
                cv_B = Now() & " <" & cv_LogLevel & "> " & cv_Record
                cv_previousText = ""
                cv_MaxLineInInfoText = Int(Info.Height / 12)
                If Info.TextBox1.LineCount >= cv_MaxLineInInfoText Then
                    cv_array03() = Split(Info.TextBox1.value, Chr(10))
                    For cv_iii = 0 To cv_MaxLineInInfoText - 2
                        cv_previousText = cv_previousText & cv_array03(cv_iii) & Chr(10)
                    Next
                    Info.TextBox1.Text = cv_B & Chr(10) & cv_previousText
                    Info.TextBox1.SelStart = 0
                Else
                    cv_previousText = Info.TextBox1.Text
                    Info.TextBox1.Text = cv_B & Chr(10) & cv_previousText
                    Info.TextBox1.SelStart = 0
                End If
                Exit For
            End If
        Next
    End If
End Sub

Public Sub cs_ShowJobStatus(Optional ShowStatusInWhichCell As String = "I6", Optional w0d1 As Integer = 0)
'1程序功能：
'1   在显示屏页指定单元格显示绿色/Done， 红色/Working
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_PageName02 As String
    Dim cv_ShowStatusInWhichCell As String
    Dim cv_w0d1 As Integer
'2赋值
    cv_PageName02 = cs_FV("ScreenSheet")
    cv_ShowStatusInWhichCell = ShowStatusInWhichCell
    cv_w0d1 = w0d1
'2代码
    With ThisWorkbook.Sheets(cv_PageName02).Range(cv_ShowStatusInWhichCell)
        If cv_w0d1 = 0 Then
            .value = "Working"
            .Interior.Color = RGB(207, 1, 37)
            .Font.Color = RGB(255, 255, 255)
        Else
            .value = "Done"
            .Interior.Color = RGB(42, 167, 75)
            .Font.Color = RGB(255, 255, 255)
        End If
    End With
End Sub

Public Function cs_FindWnd(ByVal wName As String) As Long
'1程序功能：
'1   根据给出的窗口名字wName得到对应的窗体句柄号
'1程序版本：
'1  1.1
'1版本修订：
'1  1.0 >>> 原始版本
'1  1.1 >>> 删除无用代码
'2定义
    Dim cv_wName As String
'2赋值
    cv_wName = wName
'2代码
    On Error GoTo error1
    cs_FindWnd = FindWindow(vbNullString, cv_wName)
    Call cs_Log(Now() & " FindWnd，" & "窗体： " & cv_wName & " 句柄:  " & cs_FindWnd, "Debug")
    Exit Function
error1:
    Call cs_Log(Now() & " 试图发现 窗体：" & cv_wName & " 的句柄 但FindWnd出问题了", "Error")
End Function

Public Sub cs_FindPlugin()
'1程序功能：
'1   将子目录中具有plugin功能的xlsm文件探测出，并在core_plugin中显示
'1程序版本：
'1  1.0
'1版本修订：
'1  1.0 >>> 原始版本
'2定义
    Dim cv_f As String
    Dim cv_p As String
    Dim cv_d
    Dim cv_ar
    Dim cv_i As Integer
'2赋值
    Set cv_d = CreateObject("Scripting.Dictionary")
    cv_p = ThisWorkbook.Path & cs_FV("Folder For Plugin")
    cv_f = Dir(cv_p & "Plugin*.xlsm")
'2代码
    ThisWorkbook.Sheets(cs_FV("Sheet For Plugin")).Range("D5:E100").ClearContents
    Do While Len(cv_f)
        cv_d(cv_f) = ""
        cv_f = Dir
    Loop
    cv_ar = cv_d.keys
    If cv_d.Count > 0 Then
        For cv_i = 1 To cv_d.Count
            ThisWorkbook.Sheets(cs_FV("Sheet For Plugin")).Range("D" & (cv_i + 4)).value = cv_ar(cv_i - 1)
        Next
    End If
End Sub

Public Function cs_GenPluginNumber() As Long
'1程序功能：
'1   插件序列号生成器
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_ExistingValue As Long
'2赋值
'2代码
    cv_ExistingValue = cs_FV("Plugin No. :", cs_FV("Sheet For Plugin"), "F", "G")
    Call cs_WV("Plugin No. :", (cv_ExistingValue + 1), cs_FV("Sheet For Plugin"), "F", "G")
    cs_GenPluginNumber = cv_ExistingValue + 1
End Function

Public Sub cs_LoadPlugin()
'1程序功能：
'1   导入Plugin
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本

Dim cv_i As Integer
Dim cv_ii As Integer
Dim cv_iii As Integer
Dim cv_A As String
Dim cv_B As String
Dim cv_MingganZi01 As String
Dim cv_MingganZi02 As String
Dim cv_MingganZi03 As String
Dim cv_MingganZi04 As String
Dim cv_PluginName As String
Dim cv_PluginAddress As String
Dim cv_PluginNumber As String
Dim cv_PluginXLSM As Workbook
Dim cv_SheetForPlugin As String
Dim cv_sh
Dim cv_NextLoadPluginRow As Integer
Dim cv_ovbproj As VBIDE.VBProject
Dim cv_ovbcomp As VBIDE.VBComponent
Dim cv_ovbcompUC As VBIDE.VBComponent
Dim cv_ocodemod As VBIDE.CodeModule
Dim cv_sframe As String
Dim cv_lLinestart As Long
Dim cv_wb As Workbook
Dim cv_sfunc As String
Dim cv_lindex As Long
Dim cv_length As Integer

cv_SheetForPlugin = cs_FV("Sheet For Plugin")
For cv_i = 5 To 100
    If ThisWorkbook.Sheets(cv_SheetForPlugin).Range("E" & cv_i) = "Y" Then
        cv_PluginName = ThisWorkbook.Sheets(cv_SheetForPlugin).Range("D" & cv_i)
'2确保已经加载的插件不会被再次加载\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        For cv_ii = 5 To 100
            If ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I" & cv_ii).value = cv_PluginName Then
                MsgBox "Seems Plugin: " & cv_PluginName & " is loaded. Please check it at first."
                GoTo out
            End If
        Next
 '2////////////////////////////////////////////////////////////////////////////
        cv_PluginAddress = cs_FV("Excel Path") & cs_FV("Folder For Plugin") & cv_PluginName
        cv_PluginNumber = cs_GenPluginNumber()
'2导入插件中开头为s_uc的sheet（包括vba代码）\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Set cv_PluginXLSM = Workbooks.Open(cv_PluginAddress)
        For Each cv_sh In cv_PluginXLSM.Sheets
            If Left(cv_sh.name, 4) = "s_uc" Then
                cv_PluginXLSM.Sheets(cv_sh.name).Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(cv_sh.name).name = "PI" & cv_PluginNumber & "_" & cv_sh.name
'2在core_homepage中增加每个sheet的按钮
                Call cs_CreateShape("PI" & cv_PluginNumber & "_" & cv_sh.name)
            End If
        Next
'2////////////////////////////////////////////////////////////////////////////
'2导入插件的vba模块\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'2复制模块 前缀加入 “PI插件号_”
        For Each cv_ovbcomp In cv_PluginXLSM.VBProject.VBComponents
            Select Case cv_ovbcomp.type
                Case vbext_ct_StdModule
                    Set cv_ovbcompUC = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
                    cv_ovbcompUC.name = "PI" & cv_PluginNumber & "_" & cv_ovbcomp.name
'2导出插件模块中代码
                    cv_sfunc = ""
                    With cv_PluginXLSM.VBProject.VBComponents(cv_ovbcomp.name).CodeModule
                        For cv_lindex = 1 To .CountOfLines
                            cv_A = .Lines(cv_lindex, 1)
'2将插件的模块的代码中提到的 s_uc_ and f_uc_ 加上前缀 PIXXX_
                            cv_MingganZi01 = "s_uc_"
                            cv_MingganZi02 = "PI" & cv_PluginNumber & "_" & "s_uc_"
                            cv_MingganZi03 = "f_uc_"
                            cv_MingganZi04 = "PI" & cv_PluginNumber & "_" & "f_uc_"
                            cv_A = Replace(cv_A, cv_MingganZi01, cv_MingganZi02)
                            cv_A = Replace(cv_A, cv_MingganZi03, cv_MingganZi04)
                            cv_sfunc = cv_sfunc & cv_A & Chr(10)
                        Next
                    End With
'2写入代码到ACAK Core 模块中
                    With cv_ovbcompUC.CodeModule
                        cv_lLinestart = .CountOfLines + 1
                        .InsertLines cv_lLinestart, cv_sfunc
                    End With
'2将插件中的if 与 action 连到f_core_020_ifcamp 与f_040_actionscamp
                    If cv_ovbcomp.name = "f_uc_020_ifcamp" Then
                        With cv_PluginXLSM.VBProject.VBComponents(cv_ovbcomp.name).CodeModule
                            For cv_lindex = 1 To .CountOfLines
                                cv_A = .Lines(cv_lindex, 1)
                                cv_B = .Lines(cv_lindex + 2, 1)
                                If InStr(cv_A, "Public Sub") > 0 Then
                                    With ThisWorkbook.VBProject.VBComponents("func_core_020_ifcamp").CodeModule
                                        cv_lLinestart = .CountOfLines + 1
                                        .InsertLines cv_lLinestart, cs_ucSubFuncToACAK(cv_A, Val(cv_PluginNumber), cv_B, cv_ovbcomp.name)
                                    End With
                                End If
                            Next
                        End With
                    ElseIf cv_ovbcomp.name = "f_uc_040_actionscamp" Then
                         With cv_PluginXLSM.VBProject.VBComponents(cv_ovbcomp.name).CodeModule
                            For cv_lindex = 1 To .CountOfLines
                                cv_A = .Lines(cv_lindex, 1)
                                cv_B = .Lines(cv_lindex + 2, 1)
                                If InStr(cv_A, "Public Sub") > 0 Then
                                    With ThisWorkbook.VBProject.VBComponents("func_core_040_actionscamp").CodeModule
                                        cv_lLinestart = .CountOfLines + 1
                                        .InsertLines cv_lLinestart, cs_ucSubFuncToACAK(cv_A, Val(cv_PluginNumber), cv_B, cv_ovbcomp.name)
                                    End With
                                End If
                            Next
                        End With
                    End If
            End Select
        Next
'2删插件的module
        For Each cv_ovbcomp In ThisWorkbook.VBProject.VBComponents
            Select Case cv_ovbcomp.type
                Case vbext_ct_StdModule 'case 1
                    If Left(cv_ovbcomp.name, 13 + cv_length) = "PI" & CStr(cv_PluginNumber) & "_func_debug" Then
                        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(cv_ovbcomp.name)
                    End If
            End Select
        Next

'2关闭插件

        cv_PluginXLSM.Save
        cv_PluginXLSM.Close
'2注册插件到ACAK
        cv_NextLoadPluginRow = ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I1000").End(xlUp).Row + 1
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I" & cv_NextLoadPluginRow).value = cv_PluginName
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("J" & cv_NextLoadPluginRow).value = cv_PluginNumber
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("K" & cv_NextLoadPluginRow).value = ""
'2执行s_uc_plugin_setup页中“After Loaded”代码
''2先把代码PIXXX化
'        With Sheets("PI" & CStr(cv_PluginNumber) & "_s_uc_plugin_setup")
'            For cv_iii = 2 To .Range("D1000").End(xlUp).Row
'                If .Range("D" & cv_iii).value <> "" Then
'                    .Range("D" & cv_iii).value = "PI" & CStr(cv_PluginNumber) & "_" & .Range("D" & cv_iii).value
'                End If
'            Next
'        End With
'2执行
'Call cs_TakeAction("PI" & CStr(cv_PluginNumber) & "_s_uc_plugin_setup", "D")
    End If
Next
'2刷新下action与IF list,新建相关文件夹
Call cs_FindActions
Call cs_FindIFs
Call a_CheckFolder
Call a_ShowPage("core_plugin")

MsgBox "Done"
out:
Call a_ShowPage("core_plugin")
End Sub

Public Sub cs_unLoadPlugin()
'程序功能：
'   移除Plugin
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
Dim cv_ovbcomp As VBIDE.VBComponent
Dim cv_SheetForPlugin As String
Dim cv_PluginName As String
Dim cv_PluginNumber As String
Dim cv_i As Integer
Dim cv_ii As Integer
Dim cv_length As Integer
Dim cv_sh
Dim cv_lindex As Integer
Dim cv_LineStart As Integer
Dim cv_NeedfoundString As String
Dim cv_A As String

cv_SheetForPlugin = cs_FV("Sheet For Plugin")

For cv_ii = 5 To 1000
    If ThisWorkbook.Sheets(cv_SheetForPlugin).Range("K" & cv_ii) = "Y" Then
        cv_PluginName = ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I" & cv_ii).value
        cv_PluginNumber = Trim(Str(ThisWorkbook.Sheets(cv_SheetForPlugin).Range("J" & cv_ii).value))
'2删按钮
        Call cs_RemoveShape("PI" & cv_PluginNumber)
'2删sheet
        cv_length = Len(CStr(cv_PluginNumber))
        For Each cv_sh In ThisWorkbook.Sheets
            If Left(cv_sh.name, 2 + cv_length) = "PI" & cv_PluginNumber Then
                Application.DisplayAlerts = False
                ThisWorkbook.Sheets(cv_sh.name).Delete
                Application.DisplayAlerts = True
            End If
        Next
'2删插件的module
        For Each cv_ovbcomp In ThisWorkbook.VBProject.VBComponents
            Select Case cv_ovbcomp.type
                Case vbext_ct_StdModule 'case 1
                    If Left(cv_ovbcomp.name, 2 + cv_length) = "PI" & cv_PluginNumber Then
                        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(cv_ovbcomp.name)
                    End If
            End Select
        Next
'2删if 与 action中连接过来的代码
        With ThisWorkbook.VBProject.VBComponents("func_core_020_ifcamp").CodeModule
huiqu:
            For cv_lindex = 1 To .CountOfLines
                cv_A = .Lines(cv_lindex, 1)
                cv_NeedfoundString = "Public Sub " & "PI" & cv_PluginNumber
                If InStr(cv_A, cv_NeedfoundString) > 0 Then
                    cv_LineStart = cv_lindex
                    For cv_i = 0 To 6
                        .DeleteLines cv_LineStart
                    Next
                    GoTo huiqu
                End If
                
            Next
        End With
        
        With ThisWorkbook.VBProject.VBComponents("func_core_040_actionscamp").CodeModule
huiqu2:
            For cv_lindex = 1 To .CountOfLines
                cv_A = .Lines(cv_lindex, 1)
                cv_NeedfoundString = "Public Sub " & "PI" & cv_PluginNumber
                If InStr(cv_A, cv_NeedfoundString) > 0 Then
                    cv_LineStart = cv_lindex
                    For cv_i = 0 To 5
                        .DeleteLines cv_LineStart
                    Next
                    GoTo huiqu2
                End If
            Next
        End With
'2删按钮
        Call cs_RemoveShape("PI" & cv_PluginNumber)
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I" & cv_ii).value = ""
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("J" & cv_ii).value = ""
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("K" & cv_ii).value = ""
    End If
Next

'刷新下action and IF list
Call cs_FindActions
Call cs_FindIFs
Call a_ShowPage("core_plugin")
MsgBox "Done. *** The used actions in M02 and used if in core_if should be removed manually. "
End Sub

Public Function cs_ucSubFuncToACAK(ucSubFuncSentence As String, PluginNumber As Long, Remark As String, PIXXXModuleName As String)
'1程序功能：
'1   将插件中的目标插件转化为ACAK可以引用的sub 和 function, 目前只适用于SUB
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_ucSubFuncSentence As String
    Dim cv_SubFunction As String
    Dim cv_SubFuncName  As String
    Dim cv_NewSubFuncName  As String
    Dim cv_NewArguments As String
    Dim cv_PluginNumber As Long
    Dim cv_Remark As String
    Dim cv_PIXXXModuleName As String
    Dim cv_aaa As String
    Dim cv_array01() As String
    Dim cv_d
    Dim cv_ParenthesesStart As Long
    Dim cv_ParenthesesEnd As Long
    Dim cv_SubFuncAllArguments As String
    Dim cv_i As Long
    Dim cv_iii As Long
    Dim cv_n As Long
    Dim cv_Sentence01 As String
    Dim cv_Sentence02 As String
    Dim cv_Sentence03 As String
    Dim cv_Sentence04 As String
    Dim cv_Sentence05 As String
    Dim cv_Sentence06 As String
    
    cv_ucSubFuncSentence = ucSubFuncSentence
    cv_PluginNumber = PluginNumber
    cv_Remark = Remark
    cv_PIXXXModuleName = PIXXXModuleName
    Set cv_d = CreateObject("Scripting.Dictionary")
    
    If InStr(cv_ucSubFuncSentence, " Sub ") > 0 Then
        cv_SubFunction = "Sub"
    Else
        cv_SubFunction = "Function"
    End If
    cv_aaa = Replace(cv_ucSubFuncSentence, "Public Sub ", "")
    cv_aaa = Replace(cv_aaa, "Public Function ", "")
    cv_aaa = Replace(cv_aaa, "Private Sub ", "")
    cv_aaa = Replace(cv_aaa, "Private Function ", "")
    cv_aaa = Replace(cv_aaa, "ByRef ", "")
    cv_aaa = Replace(cv_aaa, "ByVal ", "")
    
    
    cv_ParenthesesStart = InStr(cv_aaa, "(")
    cv_ParenthesesEnd = InStr(cv_aaa, ")")
    cv_SubFuncAllArguments = Mid(cv_aaa, cv_ParenthesesStart + 1, cv_ParenthesesEnd - cv_ParenthesesStart - 1)
    cv_SubFuncName = Left(cv_aaa, cv_ParenthesesStart - 1)
    cv_array01() = Split(cv_SubFuncAllArguments, ", ")
    '得到array01 有多少行 ubound +1
    cv_iii = UBound(cv_array01)
    For cv_i = 0 To cv_iii
        cv_n = InStr(cv_array01(cv_i), " As ")
        If cv_n > 0 Then
            cv_array01(cv_i) = Left(cv_array01(cv_i), cv_n - 1)
        End If
        '-----------------------------------------------------
        cv_n = InStr(cv_array01(cv_i), "Optional")
        If cv_n > 0 Then
            cv_array01(cv_i) = Right(cv_array01(cv_i), Len(cv_array01(cv_i)) - 8 + cv_n - 1)
        End If
        '-----------------------------------------------------
        If cv_i <> cv_iii And cv_i >= 0 Then
            cv_NewArguments = cv_NewArguments & cv_array01(cv_i) & ", "
        Else
            cv_NewArguments = cv_NewArguments & cv_array01(cv_i)
        End If
    Next
    cv_NewSubFuncName = "PI" & cv_PluginNumber & "_" & cv_SubFuncName
    cv_Sentence01 = "Public Sub " & cv_NewSubFuncName & "(" & cv_SubFuncAllArguments & ")" & Chr(10)
    cv_Sentence02 = "'程序功能:" & Chr(10)
    cv_Sentence03 = cv_Remark & Chr(10)
    cv_Sentence04 = "    Call " & "PI" & cv_PluginNumber & "_" & cv_PIXXXModuleName & "." & cv_SubFuncName & "(" & cv_NewArguments & ")" & Chr(10)
    cv_Sentence05 = "End Sub" & Chr(10)
    cs_ucSubFuncToACAK = cv_Sentence01 & cv_Sentence02 & cv_Sentence03 & cv_Sentence04 & cv_Sentence05
End Function

Public Sub cs_FindActions()
'1程序功能：
'1   将"func_core_040_actionscamp"中所有action都显示在“core_actions”
'1程序版本：
'1   1.0
'1版本修订：
    Dim cv_i As Long
    Dim cv_lindex As Long
    Dim cv_A As String
    Dim cv_B As String
    Dim cv_NeedfoundString As String
    
    cv_i = 0
    
    ThisWorkbook.Sheets("core_actions").Range("A2:C1000").ClearContents
    With ThisWorkbook.VBProject.VBComponents(cs_FV("ActionsInWhichVBAModule")).CodeModule
        For cv_lindex = 1 To .CountOfLines
            cv_A = .Lines(cv_lindex, 1)
            cv_NeedfoundString = "Public Sub "
            If InStr(cv_A, cv_NeedfoundString) > 0 Then
                cv_A = Replace(cv_A, "Public Sub ", "")
                cv_B = .Lines(cv_lindex + 2, 1)
                ThisWorkbook.Sheets("core_actions").Range("A" & (cv_i + 2)) = cv_i + 1
                ThisWorkbook.Sheets("core_actions").Range("B" & (cv_i + 2)) = cv_A
                ThisWorkbook.Sheets("core_actions").Range("C" & (cv_i + 2)) = cv_B
                cv_i = cv_i + 1
            End If
        Next
    End With
End Sub

Public Sub cs_FindIFs()
'1程序功能：
'1   将"func_core_020_ifcamp"中所有action都显示在“core_ifcamp”
'1程序版本：
'1   1.0
'1版本修订：
    Dim cv_i As Long
    Dim cv_lindex As Long
    Dim cv_A As String
    Dim cv_B As String
    Dim cv_NeedfoundString As String
    cv_i = 0
    ThisWorkbook.Sheets("core_ifcamp").Range("A2:C1000").ClearContents
    With ThisWorkbook.VBProject.VBComponents(cs_FV("IFInWhichVBAModule")).CodeModule
        For cv_lindex = 1 To .CountOfLines
            cv_A = .Lines(cv_lindex, 1)
            cv_NeedfoundString = "Public Sub "
            If InStr(cv_A, cv_NeedfoundString) > 0 Then
                cv_A = Replace(cv_A, "Public Sub ", "")
                cv_B = .Lines(cv_lindex + 2, 1)
                ThisWorkbook.Sheets("core_ifcamp").Range("A" & (cv_i + 2)) = cv_i + 1
                ThisWorkbook.Sheets("core_ifcamp").Range("B" & (cv_i + 2)) = cv_A
                ThisWorkbook.Sheets("core_ifcamp").Range("C" & (cv_i + 2)) = cv_B
                cv_i = cv_i + 1
            End If
        Next
    End With
End Sub

Public Sub cs_CreateShape(shapename As String)
'1程序功能：
'1   在core_homepage中显示 被加载插件的sheet
'1程序版本：
'1   1.0
'1版本修订：
'1    1.0
    Dim cv_myshape As Shape
    Dim cv_shapename As String
    Dim cv_i As Long
    Dim cv_x As Long
    Dim cv_y As Long
    Dim cv_Widthhh As Long
    Dim cv_Heighttt As Long

    cv_shapename = shapename
    
    With ThisWorkbook.Sheets("core_homepage_setup")
        For cv_i = 1 To 24
            If .Range("G" & cv_i + 1) <> "Occupied" Then
                cv_x = .Range("B" & cv_i + 1)
                cv_y = .Range("C" & cv_i + 1)
                cv_Widthhh = .Range("D" & cv_i + 1)
                cv_Heighttt = .Range("E" & cv_i + 1)
                .Range("F" & cv_i + 1) = cv_shapename
                .Range("G" & cv_i + 1) = "Occupied"
                Set cv_myshape = ThisWorkbook.Sheets("core_homepage").Shapes.AddShape(msoShapeRound2DiagRectangle, cv_x, cv_y, cv_Widthhh, cv_Heighttt)
                cv_myshape.TextFrame.Characters.Text = cv_shapename
                cv_myshape.name = cv_shapename
                cv_myshape.OnAction = "'" & "a_ShowPage " & """" & cv_shapename & """" & "'"
                Exit For
            End If
        Next
    End With
End Sub

Public Sub cs_RemoveShape(PIAndPInumber As String)
'1程序功能：
'1   在core_homepage中删除name为shapename的按钮
'1程序版本：
'1   1.0
'1版本修订：
'1    1.0
    Dim cv_myshape As Shape
    Dim cv_myshapename As String
    Dim cv_n As Long
    Dim cv_PIAndPInumber As String
    Dim cv_i As Long
    
    cv_PIAndPInumber = PIAndPInumber
    
    For Each cv_myshape In ThisWorkbook.Sheets("core_homepage").Shapes
        cv_n = Len(cv_PIAndPInumber)
        If Left(cv_myshape.name, cv_n) = cv_PIAndPInumber Then
            cv_myshapename = cv_myshape.TextFrame.Characters.Text
            cv_myshape.Delete
            For cv_i = 1 To 24
                If ThisWorkbook.Sheets("core_homepage_setup").Range("F" & cv_i + 1) = cv_myshapename Then
                    ThisWorkbook.Sheets("core_homepage_setup").Range("G" & cv_i + 1) = ""
                    ThisWorkbook.Sheets("core_homepage_setup").Range("F" & cv_i + 1) = ""
                End If
            Next
        End If
    Next

End Sub

Public Sub cs_ShowInfoWindow()
'1程序功能：
'1   显示ErrorForm
'1程序版本：
'1   1.0
'1版本修订：
'1    1.0
    Info.Show 0
End Sub

