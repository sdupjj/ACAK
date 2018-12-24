Attribute VB_Name = "func_core_040_actionscamp"
'-----------------------------------------------
'模块功能:
'   此模块用于放置zhazhupai006中的可执行的actioncamp
' a=action
'-----------------------------------------------

Public Sub X8()
'程序功能：
'   空程序
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
End Sub


Public Sub a_AddOne(Whichcount As String)
'程序功能：
'   为count原有的值增加1
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Call cs_AddOne(Whichcount)
End Sub

Public Sub a_ReduceOne(Whichcount As String)
'程序功能：
'   为count原有的值减少1
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Call cs_RedOne(Whichcount)
End Sub

Public Sub a_BeZero(Whichcount As String)
'程序功能：
'   设置count的值为0
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Call cs_BeZero(Whichcount)
End Sub

Public Sub a_BeOne(Whichcount As String)
'程序功能：
'   设置count的值为1
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Call cs_BeOne(Whichcount)
End Sub

Public Sub a_ShapeShow(PS As String)
    '功能：
    '   显示备注，which放在图形中
    '版本：
    '   1.0
    '每一版修订：
    '   1.0 >>> 原始版本
        On Error Resume Next
        For Each ashape In ActiveSheet.Shapes
            If Left(ashape.TextFrame2.TextRange.Characters.Text, 3) = PS Then
                ashape.Visible = True
            End If
        Next
End Sub

Public Sub a_ShapeHide()
    '功能：
    '   隐藏备注
    '版本：
    '   1.0
    '每一版修订：
    '   1.0 >>> 原始版本
        On Error Resume Next
        For Each ashape In ActiveSheet.Shapes
            If Left(ashape.TextFrame2.TextRange.Characters.Text, 2) = "PS" Then
                ashape.Visible = False
            End If
        Next
End Sub
Public Sub a_OpenOne()
    '功能：
    '   设置one引擎能被打开
    '版本：
    '   1.0
    '每一版修订：
    '   1.0 >>> 原始版本
    Call cs_WV("Open Engine", "Y")
    Application.StatusBar = "...ONE已经打开"
End Sub

Public Sub a_CloseOne()
    '功能：
    '   设置one引擎不能被打开
    '版本：
    '   1.0
    '每一版修订：
    '   1.0 >>> 原始版本
    Call cs_WV("Open Engine", "N")
    Application.StatusBar = "...one已经关闭"
End Sub

Public Sub a_OpenActions()
    '功能：
    '   打开执行动作功能
    '版本：
    '   1.0
    '每一版修订：
    '   1.0 >>> 原始版本
    Call cs_WV("Actions", "Y")
    Application.StatusBar = "...动作功能已经打开"
End Sub

Public Sub a_CloseActions()
    '功能：
    '   关闭执行动作功能
    '版本：
    '   1.0
    '每一版修订：
    '   1.0 >>> 原始版本
    Call cs_WV("Actions", "N")
    Application.StatusBar = "...动作功能已经关闭"
End Sub

Public Sub a_IniExcelName()
'程序功能：
'   把excel工作簿本身的名字，写入 setup
'程序版本：
'   1.0
'   1.1
'版本修订：
'   1.0 >>> 原始版本
'   1.1 >>> 由于excel 2016 365 excel 名字会有些改变，原来的程序得不到句柄，所以需要改进代码。
    PotentialName1 = ActiveWorkbook.Name & " - Excel"
    PotentialName2 = ActiveWorkbook.Name & " - Saved"
    PotentialName3 = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & " - Excel"
    PotentialName4 = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5) & " - Saved"
    If Val(cs_FindWnd(PotentialName1)) > 0 Then
        RealName = PotentialName1
    ElseIf Val(cs_FindWnd(PotentialName2)) > 0 Then
        RealName = PotentialName2
    ElseIf Val(cs_FindWnd(PotentialName3)) > 0 Then
        RealName = PotentialName3
    ElseIf Val(cs_FindWnd(PotentialName4)) > 0 Then
        RealName = PotentialName4
    Else
        MsgBox "Contact this smart person: 15026846502"
        Exit Sub
    End If
    ExcelName = RealName
    Call cs_WV("ACAK file name", ExcelName)

End Sub

Public Sub a_SimpleExcel(Optional pro As String)
'程序功能：
'   是否简单化excelsheet页的显示 让他看上去更像一个程序
'程序版本：
'   1.0
'   1.1
'版本修订：
'   1.0 >>> 原始版本
'   1.1 >>> 新版本，将控制值放在单独页面
TrueOrFalse = cs_FV("displaymode", "core_display_setup", "C", "D")
If pro = "pro" Then
    TrueOrFalse = "pro"
End If
Select Case TrueOrFalse
    Case "pro"
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        Application.DisplayFormulaBar = True
        Application.DisplayStatusBar = True
        ActiveWindow.DisplayWorkbookTabs = True
        ActiveWindow.DisplayHeadings = True
        ActiveWindow.DisplayHorizontalScrollBar = True
        ActiveWindow.DisplayVerticalScrollBar = True
    Case "simple"
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
        Application.DisplayFormulaBar = False
        Application.DisplayStatusBar = False
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayHorizontalScrollBar = False
        ActiveWindow.DisplayVerticalScrollBar = False
End Select
End Sub

Public Sub a_ShowExcelSize(MinNormalMax As String)
'程序功能：
'   最大化，最小化，正常化excel工具
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Select Case MinNormalMax
        Case "Min"
            Application.WindowState = xlMinimized
        Case "Normal"
            Application.WindowState = xlNormal
        Case "Max"
            Application.WindowState = xlMaximized
    End Select
End Sub

Public Sub a_ShowPage(PageName As String)
'程序功能：
'   显示指定的sheet,同时保留welcomepage和homepage，其他页隐藏
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
For i = 1 To Sheets.Count
    If Sheets(i).Name <> "core_screen" And Sheets(i).Name <> "core_homepage" Then
        Sheets(i).Visible = 0
    Else
        Sheets(i).Visible = 1
    End If
Next
For i = 1 To Sheets.Count
    If Sheets(i).Name = PageName Then
        Sheets(i).Visible = 1
        Sheets(i).Select
        Call a_SimpleExcel   '打开新窗口时候，检查下要求窗口的样式
    End If
Next
End Sub

Public Sub a_CheckFolder()
''程序功能：
''   检查指定的文件夹是否存在，不存在则创建
''程序版本：
''   1.0
''版本修订：
''   1.0 >>> 原始版本
    Dim excelpath As String
    excelpath = ThisWorkbook.path
    folder01 = cs_FV("Logfolder")
    folderpath01 = excelpath & folder01
    If Dir(folderpath01, vbDirectory) = "" Then MkDir folderpath01
    
    folder02 = cs_FV("picfolder")
    folderpath02 = excelpath & folder02
    If Dir(folderpath02, vbDirectory) = "" Then MkDir folderpath02
End Sub

Public Sub a_LogInTXT()
'程序功能：
'   将log 页里面的内容加入到TXT中，log页清空
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Dim excelpath As String
    excelpath = ThisWorkbook.path
    folder = cs_FV("Logfolder")
    Dim logpath As String
    Dim logname As String
    Dim PageName As String
    Dim RowE As Long
    logpath = excelpath & folder
    logname = Application.Text(Now(), "yyyymmddhhmmss") & ".txt"
    logpath = logpath & logname
    PageName = cs_FV("LogSheet")
    RowE = Sheets(PageName).Range("B150000").End(xlUp).Row
    If RowE > 1 Then
        Open logpath For Output As #1
        For i = 1 To RowE
            S = Sheets(PageName).Range("B" & i).Value
            Print #1, S
            Sheets(PageName).Range("B" & i).Value = ""
        Next
        Close #1
        Range("B" & 2 & ":B" & RowE + 1).ClearContents
    End If
End Sub

Public Sub a_initialM00()
'程序功能：
'   将M00 中的值归0
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
Dim i As Long
Dim ColN As Long
Dim ColS As Long
Dim Row2M0 As Long
Dim Row4M0 As Long
Dim PageName02 As String
PageName02 = cs_FV("ScreenSheet")
ColN = cs_FV("M01 Cols")
ColS = cs_FV("M01 Col Start Number")
Row2M0 = cs_FV("M00 ROW TOTAL") '记录M00， 单列全局可执行次数，所在行
Row4M0 = cs_FV("M00 ROW ONE") '记录M00， 单列单次可执行次数，所在行
For i = ColS To ColS + ColN - 1
    Sheets(PageName02).Cells(Row4M0, i).Value = 0
Next
For i = ColS To ColS + ColN - 1
    Sheets(PageName02).Cells(Row2M0, i).Value = 0
Next
End Sub

Public Sub a_changeOneLoppNumber(loopnumber As Long)
'程序功能：
'   设置one引擎的运作次数
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
Call cs_WV("Engine Loop", loopnumber)
End Sub



Public Sub a_MouseMove(ByVal x As Long, ByVal y As Long, ByVal ttime As Long)
Call MouseMove(x, y, ttime)
End Sub
Public Sub a_MouseClick(ByVal L0R1 As String, ByVal ttime As Long)
Call MouseClick(L0R1, ttime)
End Sub
Public Sub a_SendKeyss(keys As String)
Call SendKeyss(keys)
End Sub
