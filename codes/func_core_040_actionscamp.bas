Attribute VB_Name = "func_core_040_actionscamp"
'-----------------------------------------------
'   All available actions
'   a=action
'-----------------------------------------------
Option Explicit

Public Sub a_X8()
'1程序功能：
'1   空程序
'1程序版本：
'1  1.0
'版本修订：
'1   1.0 >>> 原始版本
End Sub

Public Sub a_AddOne(Whichcount As String)
'1程序功能：
'1   为count原有的值增加1
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_WhichCount As String
'2赋值
    cv_WhichCount = Whichcount
'2代码
    Call cs_AddOne(cv_WhichCount)
End Sub

Public Sub a_ReduceOne(Whichcount As String)
'1程序功能：
'1   为count原有的值减少1
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
    Dim cv_WhichCount As String
'2赋值
    cv_WhichCount = Whichcount
'2代码
    Call cs_ReduceOne(cv_WhichCount)
End Sub

Public Sub a_BeZero(Whichcount As String)
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
    Call cs_BeZero(cv_WhichCount)
End Sub

Public Sub a_BeOne(Whichcount As String)
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
    Call cs_BeOne(cv_WhichCount)
End Sub

Public Sub a_ShapeShow(PS As String)
'1功能：
'1   显示备注，which放在图形中
'版本：
'1   1.0
'1每一版修订：
'1   1.0 >>> 原始版本
        On Error Resume Next
        Dim cv_PS As String
        Dim cv_ashape
        cv_PS = PS
        For Each cv_ashape In ActiveSheet.Shapes
            If Left(cv_ashape.TextFrame.Characters.Text, 3) = cv_PS Then
                cv_ashape.Visible = True
            End If
        Next
End Sub

Public Sub a_ShapeHide()
'1功能：
'1   隐藏备注
'1版本：
'1   1.0
'1每一版修订：
'1   1.0 >>> 原始版本
        On Error Resume Next
        Dim cv_ashape
        For Each cv_ashape In ActiveSheet.Shapes
            If Left(cv_ashape.TextFrame.Characters.Text, 2) = "PS" Then
                cv_ashape.Visible = False
            End If
        Next
        
'2如果由于备注太多导致运行过慢。。需要摧毁使用以下代码
'    Dim cv_myshape As Shape
'    Dim cv_myshapename As String
'    Dim cv_n As Integer
'    Dim cv_PIAndPInumber As String
'    Dim cv_i As Integer
'
'    For Each cv_myshape In ActiveSheet.Shapes
'        If Left(cv_myshape.TextFrame2.TextRange.Characters.Text, 3) = "PS2" Then
'            cv_myshape.Delete
'        End If
'    Next

End Sub
Public Sub a_OpenOne()
'1功能：
'1   设置one引擎能被打开
'1版本：
'1   1.0
'1每一版修订：
'1   1.0 >>> 原始版本
    Call cs_WV("Open Engine", "Y")
    Call cs_Log("One引擎已经打开", "Info")
End Sub

Public Sub a_CloseOne()
Attribute a_CloseOne.VB_ProcData.VB_Invoke_Func = "S\n14"
'1功能：
'1   设置one引擎不能被打开
'1版本：
'1   1.0
'1每一版修订：
'1   1.0 >>> 原始版本
    Call cs_WV("Open Engine", "N")
    Call cs_Log("One引擎已经关闭", "Info")
End Sub

Public Sub a_OpenActions()
'1功能：
'1   打开执行动作功能
'1版本：
'1   1.0
'1每一版修订：
'1   1.0 >>> 原始版本
    Call cs_WV("Actions", "Y")
    Call cs_Log("One引擎执行动作功能已经打开", "Info")
End Sub

Public Sub a_CloseActions()
'1功能：
'1   关闭执行动作功能
'1版本：
'1   1.0
'1每一版修订：
'1   1.0 >>> 原始版本
    Call cs_WV("Actions", "N")
    Call cs_Log("One引擎执行动作功能已经关闭", "Info")
End Sub

Public Sub a_IniExcelName()
'1程序功能：
'1   把excel工作簿本身的名字与ACAK的文档路径，写入 setup
'1程序版本：
'1   1.3
'1版本修订：
'1   1.0 >>> 原始版本
'1   1.1 >>> 由于excel 2016 365 excel 名字会有些改变，原来的程序得不到句柄，所以需要改进代码。
'1   1.2 >>> 增加一个新功能，自动将ACAK所在的文件夹位置更新到“core_setup”页里面的Excel Path变量
'1   1.3 >>> 简化程序
    Dim cv_PotentialName1 As String
    Dim cv_RealName As String
    Dim cv_ExcelName As String
    Dim cv_excelpath As String
    
    cv_PotentialName1 = Application.Caption
    If Val(cs_FindWnd(cv_PotentialName1)) > 0 Then
        cv_RealName = cv_PotentialName1
    Else
        MsgBox "Contact this smart person: sdupjj1987@163.com"
        Exit Sub
    End If
    cv_ExcelName = cv_RealName
    Call cs_WV("ACAK file name", cv_ExcelName)
    
    cv_excelpath = ThisWorkbook.Path
    Call cs_WV("Excel Path", cv_excelpath)
End Sub

Public Sub a_SimpleExcel(Optional pro As String)
'1程序功能：
'1   是否简单化excelsheet页的显示 让他看上去更像一个程序
'1程序版本：
'1   1.1
'1版本修订：
'1   1.0 >>> 原始版本
'1   1.1 >>> 新版本，将控制值放在单独页面
    Dim cv_TrueOrFalse As String
    Dim cv_pro As String
    
    cv_pro = pro
    cv_TrueOrFalse = cs_FV("displaymode")
    
    If cv_pro = "pro" Then
        cv_TrueOrFalse = "pro"
    End If
    Select Case cv_TrueOrFalse
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
'1程序功能：
'1   最大化，最小化，正常化excel工具
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
    Dim cv_MinNormalMax As String
    cv_MinNormalMax = MinNormalMax
    Select Case cv_MinNormalMax
        Case "Min"
            Application.WindowState = xlMinimized
        Case "Normal"
            Application.WindowState = xlNormal
        Case "Max"
            Application.WindowState = xlMaximized
    End Select
End Sub

Public Sub a_ShowPage(pagename As String)
'1程序功能：
'1   显示指定的sheet,同时保留welcomepage和homepage，其他页隐藏
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
On Error Resume Next
    Dim cv_PageName As String
    Dim cv_i As Integer
    Dim cv_sh As Worksheet
    
    cv_PageName = pagename
    For Each cv_sh In ThisWorkbook.Worksheets
        If cv_sh.name <> cs_FV("ScreenSheet") And cv_sh.name <> "core_homepage" Then
            If cs_FV("HideOtherPagesWhenActiveAPage") = "Y" Then
                cv_sh.Visible = 0
            End If
        Else
            cv_sh.Visible = 1
        End If
    Next
    For Each cv_sh In ThisWorkbook.Worksheets
        If cv_sh.name = cv_PageName Then
            cv_sh.Visible = 1
            cv_sh.Select
            Call a_SimpleExcel   '打开新窗口时候，检查下要求窗口的样式
            Exit For
        End If
    Next
End Sub

Public Sub a_CheckFolder()
'1程序功能：
'1   检查在setup页,指定的文件夹是否存在，不存在则创建
'1程序版本：
'1   1.1
'1版本修订：
'1   1.0 >>> 原始版本
'1   1.1 >>> 需要动态根据setup页中出现的folders 来缺点哪些文件夹需要新增
    Dim cv_excelpath As String
    Dim cv_vname As String
    Dim cv_n As Integer
    Dim cv_i As Integer
    Dim cv_folder01 As String
    Dim cv_folderpath01 As String
    Dim cv_sh
    cv_excelpath = ThisWorkbook.Path
'检查原生 setup页 要不要生产文件夹
    For cv_i = 1 To 1000
        cv_vname = ThisWorkbook.Sheets("core_setup").Range("F" & cv_i).value
        cv_n = InStr(cv_vname, " ")
        If cv_n > 1 Then
            If Left(cv_vname, cv_n - 1) = "Folder" Then
                cv_folder01 = cs_FV(cv_vname)
                cv_folderpath01 = cv_excelpath & cv_folder01
                If Dir(cv_folderpath01, vbDirectory) = "" Then MkDir cv_folderpath01
            End If
        End If
    Next
End Sub

Public Sub a_LogInTXT()
'1程序功能：
'1   将log 页里面的内容加入到TXT中，log页清空
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
    Dim cv_excelpath As String
    Dim cv_logpath As String
    Dim cv_logname As String
    Dim cv_PageName As String
    Dim cv_folder As String
    Dim cv_LogInArray() As Variant
    Dim cv_RowE As Long
    Dim cv_s As String
    Dim cv_i As Long
    cv_excelpath = ThisWorkbook.Path
    cv_folder = cs_FV("Folder For log")
    cv_logpath = cv_excelpath & cv_folder
    cv_logname = Application.Text(Now(), "yyyymmddhhmmss") & ".txt"
    cv_logpath = cv_logpath & cv_logname
    cv_PageName = cs_FV("LogSheet")
    cv_RowE = ThisWorkbook.Sheets(cv_PageName).Range("B1000000").End(xlUp).Row
    If cv_RowE = 1 Then
        GoTo A
    End If
    cv_LogInArray = ThisWorkbook.Sheets(cv_PageName).Range("B1:B" & cv_RowE).value
    If cv_RowE > 1 Then
        Open cv_logpath For Output As #1
        For cv_i = 1 To cv_RowE
            cv_s = cv_LogInArray(cv_i, 1)
            Print #1, cv_s
        Next
        Close #1
         ThisWorkbook.Sheets(cv_PageName).Range("B" & 2 & ":B" & cv_RowE + 1).Delete Shift:=xlUp
    End If
A:
    Call cs_BeZero("count03")
End Sub

Public Sub a_initialM00()
'1程序功能：
'1   将M00 中的值归0
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
    Dim cv_i As Long
    Dim cv_ColN As Long
    Dim cv_ColS As Long
    Dim cv_Row2M0 As Long
    Dim cv_Row4M0 As Long
    Dim cv_PageName02 As String
    cv_PageName02 = cs_FV("ScreenSheet")
    cv_ColN = cs_FV("M01 Cols")
    cv_ColS = cs_FV("M01 Col Start Number")
    cv_Row2M0 = cs_FV("M00 ROW TOTAL") '记录M00， 单列全局可执行次数，所在行
    cv_Row4M0 = cs_FV("M00 ROW ONE") '记录M00， 单列单次可执行次数，所在行
    For cv_i = cv_ColS To cv_ColS + cv_ColN - 1
        ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row4M0, cv_i).value = 0
    Next
    For cv_i = cv_ColS To cv_ColS + cv_ColN - 1
        ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row2M0, cv_i).value = 0
    Next
End Sub

Public Sub a_changeOneLoppNumber(loopnumber As Long)
'1程序功能：
'1   设置one引擎的运作次数
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
    Dim cv_loopnumber As Long
    cv_loopnumber = loopnumber
    Call cs_WV("Engine Loop", cv_loopnumber)
    Call cs_Log("设置One引擎滚动次数: " & CStr(cv_loopnumber), "Info")
End Sub

Public Sub a_ExportCode()
'1功能：
'1   导出ACAK中所有代码
'1版本：
'1   1.0
'1每一版修订：
'1   1.0 >>> 原始版本
    Call cs_ExportCode
End Sub

Public Sub a_FindPlugin()
'1程序功能：
'1   将子目录中具有plugin功能的xlsm文件探测出，并在core_plugin中显示，并导入
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
    Call cs_FindPlugin
End Sub


Public Sub a_IFController()
'1程序功能：
'1   运行一次c_IFController, 让IF sensor 都运行一次
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
    Call c_IFController
End Sub

Public Sub a_CheckACAKCore()
'1程序功能：
'1   检查ACAKCore是否完整，页面是否都存在，引用是否都存在
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
'2定义
Dim yinyong As Variant 'yinyong 是一个2维数组，保存必须导入的引用
'yinyong = Evaluate("{" & _
'                """VBA"",""{000204EF-0000-0000-C000-000000000046}"",4,2;" & _
'                """Excel"",""{00020813-0000-0000-C000-000000000046}"",1,9;" & _
'                """stdole"",""{00020430-0000-0000-C000-000000000046}"",2,0;" & _
'                """Office"",""{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}"",2,8;" & _
'                """VBIDE"",""{0002E157-0000-0000-C000-000000000046}"",5,3;" & _
'                """MSForms"",""{0D452EE1-E08F-101A-852E-02608C4D0BB4}"",2,0" & _
'                "}")
Dim cv_i As Long
Dim cv_installed As Long
Dim cv_refed
Dim cv_sheet As Worksheet
Dim cv_sheetname As String
'2赋值
cv_i = 1
cv_installed = 0
cv_sheetname = "Core_ACAK_structure"
'2代码
'2检查指定引用有没有被导入
With Sheets(cv_sheetname)
    For cv_i = 4 To .Range("C1000").End(xlUp).Row
        cv_installed = 0
        For Each cv_refed In ThisWorkbook.VBProject.References
            If cv_refed.Guid = .Range("C" & cv_i) Then
                If cv_refed.IsBroken Then
                    ThisWorkbook.VBProject.References.Remove cv_refed
                    cv_installed = 0
                Else
                     .Range("F" & cv_i) = "Y"
                     cv_installed = 1
                     Exit For
                End If
            End If
        Next
'2如果没有导入，则导入引用
        If cv_installed = 0 Then
            Call cs_Log("Try to load reference: " & .Range("C" & cv_i), "Info")
            Call ThisWorkbook.VBProject.References.AddFromGuid(.Range("C" & cv_i), Val(.Range("D" & cv_i)), Val(.Range("E" & cv_i)))
            .Range("F" & cv_i) = "Y"
        End If
    Next
End With
'2检查sheets是否都存在
cv_i = 1
With Sheets(cv_sheetname)
    For cv_i = 4 To .Range("I1000").End(xlUp).Row
        cv_installed = 0
        For Each cv_sheet In ThisWorkbook.Worksheets
            If cv_sheet.name = .Range("I" & cv_i) Then
                cv_installed = 1
                 .Range("J" & cv_i) = "Y"
                Exit For
            End If
        Next
        If cv_installed = 0 Then
            .Range("J" & cv_i) = "N"
            Call cs_Log("Sheet " & .Range("I" & cv_i) & " can not be found. ", "Error")
        End If
    Next
End With

Exit Sub

Errhandle:
    Call cs_Log("a_CheckACAKCore met some errors, please check. ", "Error")
    Exit Sub
End Sub
