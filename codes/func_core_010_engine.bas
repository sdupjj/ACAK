Attribute VB_Name = "func_core_010_engine"
'-----------------------------------------------
'   ACAK One Engine module
'-----------------------------------------------
Option Explicit
Public Sub c_One()
'1程序功能：
'1   程序运行引擎
'1   是否能启动 第0次一定能被运行，之后 one引擎开关关闭>单次列运转上限已经达到=未找到可以执行的列>运行次数已经达到
'1程序版本：
'1   1.1
'1版本修订：
'1   1.1 >>> 简化以及美化
'1   1.0 >>> 原始版本
'2定义
    Dim cv_AA As String
    Dim cv_x As Long
    Dim cv_y As Long
    Dim cv_i As Long
    Dim cv_n As Long
    Dim cv_m As Long
    Dim cv_Rows As Long
    Dim cv_RowSS As Long
    Dim cv_RowE As Long
    Dim cv_Vrow As Long
    Dim cv_PageName02 As String  'screen sheet
    Dim cv_CellDoneWorking As String
    Dim cv_PageName03 As String  'engine setup sheet
    Dim cv_ActionsCamp As String  'actions in which vba module
    Dim cv_EngYN As String
    Dim cv_BehYN As String
    Dim cv_EngLoopNumber As Long
    Dim cv_T05
    Dim cv_T06
    Dim cv_T07
'2赋值
    cv_Vrow = cs_FV("VinScreenRow")
    cv_PageName02 = cs_FV("ScreenSheet")
    cv_PageName03 = cs_FV("EngineSetupSheet")
    cv_ActionsCamp = cs_FV("ActionsInWhichVBAModule")
    cv_Rows = cs_FV("M02 Row Start Number")
    cv_RowSS = cs_FV("M02 Rows") '得到有多少指令需要运行
    cv_EngLoopNumber = cs_FV("Engine Loop")
    cv_CellDoneWorking = cs_FV("ShowStatusInWhichCellInScreenSheet")
'2代码
    Call cs_ShowJobStatus(cv_CellDoneWorking, 0)
    For cv_n = 0 To cv_EngLoopNumber + 1
        Select Case cv_n
            '2 One第0次运行，初始化运行
            Case 0
                Call cs_TakeAction(cv_PageName03, "D")
             '2 One 第1次至规定次数运行
            Case 1 To cv_EngLoopNumber
                '2 检查引擎是否很启动
                cv_EngYN = cs_FV("Open Engine")
                '2 检查引擎能否允许执行动作
                cv_BehYN = cs_FV("Actions")
                If cv_EngYN = "Y" Then
                    '2 计数 ACAK打开后 One所有运行次数
                    Call cs_AddOne("count01")
                    '2 计数 One重新发动后，运行次数
                    Call cs_AddOne("count02")
                    cv_T07 = Timer
                    '2 运行每个IF<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    Call c_IFController
                    '2 "V"标识 所在行清空
                    ThisWorkbook.Sheets(cv_PageName02).Range("G" & cv_Vrow & ":AG" & cv_Vrow).ClearContents
                    '2 找出“V”应该在那一列，标出，并返回列号（数字）
                    cv_x = c_GetRightColInM01()
                    DoEvents
                    '2 根据提供的列号，判断是否已经超过最大运行次数 ，可以运行则返回 1， 不可以运行则返回0
                    cv_y = c_IfRightColInM01ShouldRun(cv_x)
                    If cv_x > 0 And cv_y > 0 And cv_BehYN = "Y" Then '可以运行
                            Call cs_TakeAction(cv_PageName02, cs_num2asc2(cv_x), cv_Rows, cv_Rows + cv_RowSS - 1, "Y")
                            Call cs_Log("One Engine Loop 耗时： " & Str(Application.Text((Timer - cv_T07), "0.0000")), "Print")
                    ElseIf cv_x > 0 And cv_y > 0 And cv_BehYN = "N" Then 'One动作被设置为不可以执行
                        Call cs_Log("ACAK被设置为：动作将不被执行。", "Info")
                    ElseIf cv_x < 0 Then '找不到对应的列'
                        Call cs_Log("M01中找不到对应的列，动作将不被执行。", "Info")
                        Exit For
                    ElseIf cv_y = 0 Then
                        Call cs_Log("M01 & M02 对应的列超过了规定循环次数，动作将不被执行。", "Info")
                        Exit For
                    End If
                Else 'One引擎被设置未不可被执行
                    Call cs_Log("One引擎已经被关闭", "Info")
                    Exit For
                End If
            Case cv_EngLoopNumber + 1
            '2 One已经正常run完毕，需要收尾
                Call cs_TakeAction(cv_PageName03, "H")
                Call cs_Log("One引擎已经达到设置的最大循环次数，不再运行。", "Info")
        End Select
    Next
    Call cs_ShowJobStatus(cv_CellDoneWorking, 1)
End Sub

Public Function c_GetRightColInM01() As Long
'1程序功能：
'1   通过对比侦测的结果，在M01中找到对应的列,并在对应的列头上标个“V"
'1程序版本：
'1   1.1
'1版本修订：
'1   1.0 >>> 原始版本,并无颜色标记
'1   1.1 >>> debug
'定义
Dim cv_x As Long
Dim cv_i As Long
Dim cv_ii As Long
Dim cv_ix As Long
Dim cv_n As Long
Dim cv_max As Long
Dim cv_IFColN As Long
Dim cv_VRowN As Long
Dim cv_RowN As Long
Dim cv_ColN As Long
Dim cv_Rows As Long
Dim cv_ColS As Long
Dim cv_JudgeWay As Long
Dim cv_Arr01() As Long
Dim cv_PageName02 As String
Dim cv_T03
'赋值
cv_IFColN = cs_FV("IFdisplayinScreenCol")
cv_VRowN = cs_FV("VinScreenRow")
cv_RowN = cs_FV("M01 Rows")
cv_ColN = cs_FV("M01 Cols")
cv_Rows = cs_FV("M01 Row Start Number")
cv_ColS = cs_FV("M01 Col Start Number")
cv_JudgeWay = cs_FV("EngineColumnSelectMethod")
cv_PageName02 = cs_FV("ScreenSheet")
'代码

cv_T03 = Timer

If cv_JudgeWay = 1 Then
    cv_x = 0
    cv_n = 0
    For cv_i = 0 To cv_ColN - 1
        cv_ix = c_CompareTwoWordInM01(cv_IFColN, cv_ColS + cv_i, cv_Rows, cv_Rows + cv_RowN - 1)(1)
        If cv_ix = 1 Then
            Sheets(cv_PageName02).Cells(cv_VRowN, cv_ColS + cv_i).value = "V"
            cv_x = 1
            Exit For
        End If
    Next
    If cv_x = 1 Then
        c_GetRightColInM01 = cv_ColS + cv_i
    Else
        c_GetRightColInM01 = -1
    End If
    
ElseIf cv_JudgeWay = 2 Then
            cv_x = 0
            cv_n = 0
            cv_max = 0
            ReDim cv_Arr01(cv_ColN - 1)
        '----------------------------------
        '确定有符合条件的列
            For cv_i = 0 To cv_ColN - 1
                cv_ix = c_CompareTwoWordInM01(cv_IFColN, cv_ColS + cv_i, cv_Rows, cv_Rows + cv_RowN - 1)(1)
                If cv_ix = 1 Then
                    cv_x = 1
                    Exit For
                End If
            Next
        '-----------------------------------
            For cv_i = 0 To cv_ColN - 1
                cv_Arr01(cv_i) = c_CompareTwoWordInM01(cv_IFColN, cv_ColS + cv_i, cv_Rows, cv_Rows + cv_RowN - 1)(2)
            Next
        '-----------------------------------
            cv_max = cv_Arr01(0)
            For cv_i = 0 To cv_ColN - 2
                If cv_max < cv_Arr01(cv_i + 1) Then
                    cv_max = cv_Arr01(cv_i + 1)
                End If
            Next
        '-----------------------------------
            For cv_i = 0 To cv_ColN - 1
                If cv_Arr01(cv_i) = cv_max Then
                    cv_ii = cv_i
                    Exit For
                End If
            Next
            If cv_x = 1 And cv_max >= 0 Then
                ThisWorkbook.Sheets(cv_PageName02).Cells(cv_VRowN, cv_ColS + cv_ii).value = "V"
                c_GetRightColInM01 = cv_ColS + cv_ii
            Else
                c_GetRightColInM01 = -1
            End If
End If
Call cs_Log("Found Right Column Total " & Str(Application.Text((Timer - cv_T03), "0.0000")), "Debug") '查看找到正确列的时间
End Function

Public Function c_CompareTwoWordInM01(WhichColNum01 As Long, WhichColNum02 As Long, U As Long, L As Long) As Long()
'1程序功能：
'1   对比矩阵01中的两列，如果是8则跳过，如果一致返回1，如果不一致返回0
'1程序版本：
'1   1.1
'1版本修订：
'1   1.0 >>> 原始版本,并无颜色标记
'1   1.1 >>> debug
    Dim cv_PageName02 As String
    Dim cv_ColNum01 As Long
    Dim cv_ColNum02 As Long
    Dim cv_A As String
    Dim cv_B As String
    Dim cv_C(1 To 2) As Long
    Dim cv_HowLong As Long
    Dim cv_i As Long
    Dim cv_u As Long
    Dim cv_l As Long
    
    cv_u = U
    cv_l = L
    cv_PageName02 = cs_FV("ScreenSheet")
    cv_ColNum01 = WhichColNum01
    cv_ColNum02 = WhichColNum02

    For cv_i = 0 To cv_l - cv_u
        If ThisWorkbook.Sheets(cv_PageName02).Cells(cv_u + cv_i, cv_ColNum01).value <> 8 And ThisWorkbook.Sheets(cv_PageName02).Cells(cv_u + cv_i, cv_ColNum02).value <> 8 Then
            cv_A = cv_A & Trim((Str(ThisWorkbook.Sheets(cv_PageName02).Cells(cv_u + cv_i, cv_ColNum01).value)))
            cv_B = cv_B & Trim((Str(ThisWorkbook.Sheets(cv_PageName02).Cells(cv_u + cv_i, cv_ColNum02).value)))
            cv_HowLong = cv_HowLong + 1
        End If
    Next
    If cv_A = cv_B Then
        cv_C(1) = 1
        cv_C(2) = cv_HowLong
    Else
        cv_C(1) = 0
        cv_C(2) = -1
    End If
    c_CompareTwoWordInM01 = cv_C
End Function

Public Function c_IfRightColInM01ShouldRun(x As Long)
'1程序功能：
'1   检查 x列 有没有超过规定的次数，没有返回1，有返回0，同时统计次数，统计次数仅针对上一次也是相同列的情况进行累加，别的都清0
'1程序版本：
'1   1.1
'1版本修订：
'1   1.0 >>> 原始版本,并无颜色标记
'1   1.1 >>> debug
Dim cv_i As Long
Dim cv_x As Long
Dim cv_quanju As Long
Dim cv_guodu As Long
Dim cv_RowN As Long
Dim cv_ColN As Long
Dim cv_Rows As Long
Dim cv_ColS As Long
Dim cv_PageName02 As String
Dim cv_Row1M0 As Long
Dim cv_Row2M0 As Long
Dim cv_Row3M0 As Long
Dim cv_Row4M0 As Long
Dim cv_M00Fuction1 As String
Dim cv_M00Fuction2 As String
Dim cv_A As Long
Dim cv_B As Long
Dim cv_a_1 As Long
Dim cv_b_1 As Long
cv_x = x
cv_A = 1
cv_B = 1
cv_quanju = 0
cv_a_1 = 1
cv_b_1 = 1
cv_RowN = cs_FV("M01 Rows")
cv_ColN = cs_FV("M01 Cols")
cv_Rows = cs_FV("M01 Row Start Number")
cv_ColS = cs_FV("M01 Col Start Number")
cv_PageName02 = cs_FV("ScreenSheet")

cv_Row1M0 = cs_FV("M00 ROW TOTAL MAX") '判断M00， 单列全局可执行次数，最大，所在行
cv_Row2M0 = cs_FV("M00 ROW TOTAL") '记录M00， 单列全局可执行次数，所在行
cv_Row3M0 = cs_FV("M00 ROW ONE MAX") '判断M00， 单列单次可执行次数，最大，所在行
cv_Row4M0 = cs_FV("M00 ROW ONE") '记录M00， 单列单次可执行次数，所在行
cv_M00Fuction1 = cs_FV("M00 Judge TOTAL MAX") ' 是否启用判断 Total max， Y or N
cv_M00Fuction2 = cs_FV("M00 Judge ONE MAX") ' 是否启用判断 ONE max， Y or N

If cv_x > 0 Then
    If cv_M00Fuction2 = "Y" Then
        If ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row4M0, cv_x).value < ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row3M0, cv_x).value Then
            cv_a_1 = 1
        Else
            cv_a_1 = 0
        End If
    End If
    If cv_M00Fuction1 = "Y" Then
        If ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row2M0, cv_x).value < ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row1M0, cv_x).value Then
            cv_b_1 = 1
        Else
            cv_b_1 = 0
        End If
    End If
    If cv_M00Fuction2 = "Y" Then
        If ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row4M0, cv_x).value < ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row3M0, cv_x).value And cv_a_1 * cv_b_1 = 1 Then
            ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row2M0, cv_x).value = ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row2M0, cv_x).value + 1
            cv_quanju = 1
            cv_guodu = ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row4M0, cv_x).value + 1
            For cv_i = cv_ColS To (cv_ColS + cv_ColN - 1)
                ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row4M0, cv_i).value = 0
            Next
            ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row4M0, cv_x).value = cv_guodu
            cv_A = 1
        Else
            cv_A = 0
        End If
    End If
    If cv_M00Fuction1 = "Y" Then
        If ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row2M0, cv_x).value < ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row1M0, cv_x).value And cv_a_1 * cv_b_1 = 1 Then
            If cv_quanju = 0 Then
                ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row2M0, cv_x).value = ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row2M0, cv_x).value + 1
            End If
            cv_B = 1
        Else
            cv_B = 0
        End If
    End If
    If cv_A * cv_B = 1 Then
        c_IfRightColInM01ShouldRun = 1
    Else
        c_IfRightColInM01ShouldRun = 0
    End If
ElseIf x < 0 Then
    c_IfRightColInM01ShouldRun = 0
End If
End Function

Public Sub c_IFController()
'1程序功能：
'1   选择 core_if 页里 judge哪一个 判断列，并执行相关接口程序（设为1的情况）
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
    Dim cv_PageName02 As String
    Dim cv_PageName03 As String
    Dim cv_PageName04 As String
    Dim cv_Range01 As String
    Dim cv_T
    Dim cv_T2
    Dim cv_i As Long
    Dim cv_AA As String
    Dim cv_CellDoneWorking As String
    
    cv_PageName02 = cs_FV("ScreenSheet")
    cv_PageName03 = cs_FV("IFSheet")
    cv_PageName04 = cs_FV("IFInWhichVBAModule")
    cv_Range01 = cs_FV("ShowStatusInWhichCellInScreenSheet")
    cv_T = Timer
    cv_CellDoneWorking = cs_FV("ShowStatusInWhichCellInScreenSheet")
    
    With ThisWorkbook.Sheets(cv_PageName03)
        For cv_i = 2 To .Range("a1000").End(xlUp).Row
            If .Range("D" & cv_i).value = 1 Then
                cv_T2 = Timer
                cv_AA = "'" & cv_PageName04 & "." & .Range("b" & cv_i) & "'"
                cs_runAA (cv_AA)
                DoEvents
                Call cs_Log(cv_AA & " " & Str(Application.Text((Timer - cv_T2), "0.0000")), "Debug")
            Else
                ThisWorkbook.Sheets(cv_PageName02).Range(.Range("A" & cv_i).value).Interior.Color = RGB(141, 145, 146)
                ThisWorkbook.Sheets(cv_PageName02).Range(.Range("A" & cv_i).value).value = 0
                DoEvents
            End If
        Next
    End With
    Call cs_Log(" IF Total : " & Str(Application.Text((Timer - cv_T), "0.0000")), "Debug")
End Sub

Public Sub c_CheckStatus()
'1程序功能：
'1   连接core_screen 中的checkstatus按钮
'1程序版本：
'1   1.0
'1版本修订：
'1   1.0 >>> 原始版本
    Dim cv_CellDoneWorking As String
    
    cv_CellDoneWorking = cs_FV("ShowStatusInWhichCellInScreenSheet")
    
    Call cs_ShowJobStatus(cv_CellDoneWorking, 0)
    Call c_IFController
    Call cs_ShowJobStatus(cv_CellDoneWorking, 1)
End Sub

