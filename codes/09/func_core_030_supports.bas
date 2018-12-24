Attribute VB_Name = "func_core_030_supports"
'-----------------------------------------------
'模块功能:
'   此模块用于放置zhazhupai006中的支持各种程序运行的一些核心小程序
'   cs=core support
'   if=interface sensor
'-----------------------------------------------
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Function cs_FV(Ipt As Variant, Optional WhichSheet As String = "core_setup", Optional WhichCol01 As String = "F", _
                Optional WhichCol02 As String = "G", Optional HowManyRows As Long = 1000) As Variant
'程序功能：
'   通过指定列中的值，找另外指定列中的值
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Dim i As Integer
    Dim x As Long
    Dim n As Integer
    With Sheets(WhichSheet)
        For x = 1 To .Range(WhichCol01 & HowManyRows).End(xlUp).Row
            If .Range(WhichCol01 & x).Value = Ipt Then
                cs_FV = .Range(WhichCol02 & x).Value
                Exit For
            End If
        Next
    End With
End Function
Public Sub cs_WV(Ipt As Variant, Wpt As Variant, Optional WhichSheet As String = "core_setup", Optional WhichCol01 As String = "F", _
                Optional WhichCol02 As String = "G", Optional HowManyRows As Long = 1000)
'程序功能：
'   通过指定列中的值Ipt，写入另外指定列中的值Wpt
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Dim i As Integer
    Dim x As Long
    Dim n As Integer

    With Sheets(WhichSheet)
        For x = 1 To .Range(WhichCol01 & HowManyRows).End(xlUp).Row
            If .Range(WhichCol01 & x).Value = Ipt Then
                    .Range(WhichCol02 & x).Value = Wpt
                Exit For
            End If
        Next
    End With
End Sub

Public Sub cs_AddOne(Whichcount As String)
'程序功能：
'   写入内容页里的 whichcount 旁边的数值加1
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Dim MiddleCount As Long
    MiddleCount = cs_FV(Whichcount, "core_count", "A", "B")
    Call cs_WV(Whichcount, MiddleCount + 1, "core_count", "A", "B")
End Sub
Public Sub cs_RedOne(Whichcount As String)
'程序功能：
'   写入内容页里的 whichcount 旁边的数值减1
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Dim MiddleCount As Long
    MiddleCount = cs_FV(Whichcount, "core_count", "A", "B")
    Call cs_WV(Whichcount, MiddleCount - 1, "core_count", "A", "B")
End Sub

Public Sub cs_BeZero(Whichcount As String)
'程序功能：
'   写入内容页里的 whichcount 旁边的数值变成0
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Call cs_WV(Whichcount, 0, "core_count", "A", "B")
End Sub

Public Sub cs_BeOne(Whichcount As String)
'程序功能：
'   写入内容页里的 whichcount 旁边的数值变成1
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Call cs_WV(Whichcount, 1, "core_count", "A", "B")
End Sub

Public Sub cs_BeValue(Whichcount As String, Value As Variant)
'程序功能：
'   写入内容页里的 whichcount 旁边的数值变成value
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Call cs_WV(Whichcount, Value, "core_count", "A", "B")
End Sub

Public Function cs_FindValue(Whichcount As String) As Variant
'程序功能：
'   发现的whichcount 旁边的数值
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    cs_FindValue = cs_FV(Whichcount, "core_count", "A", "B")
End Function


Public Sub cs_TakeAction(whichpagename As String, whichcol As String, Optional RowS As Integer = 2)
'程序功能：
'   对某页的某列动作 进行执行
'程序版本：
'   1.1
'版本修订：
'   1.0 >>> 原始版本,并无颜色标记
'   1.1 >>> 在状态栏显示动作
    Dim RowE As Integer
    Dim AA As String
    Dim i As Integer
    Dim ActionsCamp As String  'actions in which vba module
    RowE = Sheets(whichpagename).Range(whichcol & 10000).End(xlUp).Row
    ActionsCamp = cs_FV("ActionsInWhichVBAModule")
    For i = RowS To RowE
        AA = "'" & ActionsCamp & "." & Sheets(whichpagename).Range(whichcol & i).Value & "'"
        Application.StatusBar = AA
        Application.Run AA
        DoEvents
    Next
End Sub

Public Sub cs_Log(sts As Variant, level_Debug_Error_Print_Info As String)
'程序功能：
'   记录sts到core_log中
'程序版本：
'   1.0
'   1.1
'版本修订：
'   1.0 >>> 原始版本
'   1.1 >>> 增加一个参数 确定 是在什么情况下记录log
    Dim PageName As String
    Dim A As Long
    Dim B As String
    PageName = cs_FV("LogSheet")
    Call cs_AddOne("count03")
    A = cs_FindValue("count03") + 1
    B = Now() & " " & level_Debug_Error_Print_Info & " " & sts
    Sheets(PageName).Range("B" & A).Value = B
End Sub

Public Sub cs_ShowJobStatus(Optional ShowStatusInWhichCell As String = "I6", Optional w0d1 As Integer = 0)
    Dim PageName02 As String
    PageName02 = cs_FV("ScreenSheet")
    With Sheets(PageName02).Range(ShowStatusInWhichCell)
        If w0d1 = 0 Then
            .Value = "Working"
        Else
            .Value = "Done"
        End If
    End With
End Sub


Public Function cs_FindWnd(ByVal wName As String) As Long
'程序功能：
'   根据给出的窗口名字wName得到对应的窗体句柄号
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
On Error GoTo error1
    If Val(Application.Version) < 9 Then
        cs_FindWnd = FindWindow("ThunderXFrame", wName) 'XL97
    Else
        cs_FindWnd = FindWindow("ThunderDFrame", wName) 'XL2000
    End If
    If cs_FindWnd = 0 Then cs_FindWnd = FindWindow(vbNullString, wName)
    Call cs_Log(Now() & " FindWnd，" & "窗体： " & wName & " 句柄:  " & cs_FindWnd, "Debug")
    Exit Function
error1:
    Call cs_Log(Now() & " FindWnd出问题了", "Error")
End Function



