Attribute VB_Name = "func_core_010_engine"
'-----------------------------------------------
'模块功能:
'   此模块用于放置zhazhupai006中的发动机one
'   c=core
'-----------------------------------------------

Public Sub c_one()
'程序功能：
'   程序运行引擎
'   是否能启动 第0次一定能被运行，之后 one引擎开关关闭>单次列运转上限已经达到=未找到可以执行的列>运行次数已经达到
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
Dim AA As String
Dim x As Integer
Dim y As Integer
Dim i As Integer
Dim n As Long
Dim m As Long
Dim RowS As Integer
Dim RowE As Integer
Dim Vrow As Integer
Dim PageName02 As String  'screen sheet
Dim PageName03 As String  'engine setup sheet
Dim ActionsCamp As String  'actions in which vba module
Dim EngYN As String
Dim BehYN As String
Dim EngLoopNumber As Long
Vrow = cs_FV("VinScreenRow")
PageName02 = cs_FV("ScreenSheet")
PageName03 = cs_FV("EngineSetupSheet")
ActionsCamp = cs_FV("ActionsInWhichVBAModule")
RowS = cs_FV("M02 Row Start Number")
RowSS = cs_FV("M02 Rows") '得到有多少指令需要运行
EngLoopNumber = cs_FV("Engine Loop")
For n = 0 To EngLoopNumber + 1
    Select Case n
        Case 0 ' 第0次运行，初始化运行
            Call cs_TakeAction(PageName03, "D")
        Case 1 To EngLoopNumber ' 第1次运行 到 规定次数运行，正常运行
            EngYN = cs_FV("Open Engine") '看下引擎是否很启动
            BehYN = cs_FV("Actions") '看下引擎能否执行对应动作很启动
            If EngYN = "Y" Then
                Call cs_AddOne("count01")                           'ACAK打开后 one所有运行次数
                Call cs_AddOne("count02")                           'one重新发动后，运行次数
                Call c_IFController   '运行每个接口程序
                T03 = Timer
                Sheets(PageName02).Range("G" & Vrow & ":AG" & Vrow).ClearContents '"V"所在行清空
                x = c_GetRightColInM01() '找出“V”应该在那一列，标出，并返回列号
                Call cs_Log("Total " & Str((Timer - T03)), "Info") '查看找到正确列的时间
                y = c_IfRightColInM01ShouldRun(x) '根据提供的列号，判断是否已经超过单一最大运行次数
                If x > 0 And y > 0 And BehYN = "Y" Then '可以运行
                        T06 = Timer
                        For i = RowS To RowS + RowSS - 1
                            Sheets(PageName02).Range(num2asc2(x) & i).Interior.Color = 192
                            Application.StatusBar = Sheets(PageName02).Range(num2asc2(x) & i).Value
                             AA = "'" & ActionsCamp & "." & Cells(i, x) & "'"
                             T05 = Timer
                            Application.Run AA '运行程序
                            Call cs_Log(Sheets(PageName02).Range(num2asc2(x) & i).Value & " 耗时： " & Str((Timer - T05)), "Info")
                            Application.StatusBar = ""
                            Sheets(PageName02).Range(num2asc2(x) & i).Interior.Color = 16777215
                            DoEvents
                        Next
                        Call cs_Log(num2asc2(x) & "列动作 耗时： " & Str((Timer - T06)), "Info")
                        
                ElseIf x > 0 And y > 0 And BehYN = "N" Then '不可以运行
                    Application.StatusBar = "被设置为：动作将不被执行。"
                ElseIf x < 0 Or y = 0 Then
                    Application.StatusBar = "矩阵01中找不到对应的列或者对应的列超过了最大单一循环次数，没有动作将被执行。"
                    Exit For
                End If
            Else
                Application.StatusBar = "...ONE引擎已经关闭"
                Exit For
            End If
        Case EngLoopNumber + 1
        End Select
    m = m + 1
Next
If m = EngLoopNumber + 2 Then
    Application.StatusBar = "one引擎已经达到设置的最大循环次数，不再运行。"
    Call cs_TakeAction(PageName03, "H")
End If
End Sub

Public Function c_GetRightColInM01() As Integer
'程序功能：
'   通过对比侦测的结果，在matrix01中找到对应的值,并在对应的列头上标个“V"
'程序版本：
'   1.1
'版本修订：
'   1.0 >>> 原始版本,并无颜色标记
Dim x As Integer
Dim i As Integer
Dim ii As Integer
Dim ix As Integer
Dim n As Integer
Dim max As Integer
Dim IFColN As Integer
Dim VRowN As Integer
Dim RowN As Integer
Dim ColN As Integer
Dim RowS As Integer
Dim ColS As Integer
Dim JudgeWay As Integer
Dim Arr01() As Integer
Dim PageName02 As String
IFColN = cs_FV("IFdisplayinScreenCol")
VRowN = cs_FV("VinScreenRow")
RowN = cs_FV("M01 Rows")
ColN = cs_FV("M01 Cols")
RowS = cs_FV("M01 Row Start Number")
ColS = cs_FV("M01 Col Start Number")
JudgeWay = cs_FV("EngineColumnSelectMethod")
PageName02 = cs_FV("ScreenSheet")
If JudgeWay = 1 Then
    x = 0
    n = 0
    For i = 0 To ColN - 1
        ix = c_CompareTwoWordInM01(IFColN, ColS + i, RowS, RowS + RowN - 1)(1)
        If ix = 1 Then
            Sheets(PageName02).Cells(VRowN, ColS + i).Value = "V"
            x = 1
            Exit For
        End If
    Next
    If x = 1 Then
        c_GetRightColInM01 = ColS + i
    Else
        c_GetRightColInM01 = -1
    End If
    
ElseIf JudgeWay = 2 Then
            x = 0
            n = 0
            max = 0
            ReDim Arr01(ColN - 1)
        '----------------------------------
        '确定有符合条件的列
            For i = 0 To ColN - 1
                ix = c_CompareTwoWordInM01(IFColN, ColS + i, RowS, RowS + RowN - 1)(1)
                If ix = 1 Then
                    x = 1
                    Exit For
                End If
            Next
        '-----------------------------------
            For i = 0 To ColN - 1
                Arr01(i) = c_CompareTwoWordInM01(IFColN, ColS + i, RowS, RowS + RowN - 1)(2)
            Next
        '-----------------------------------
            max = Arr01(0)
            For i = 0 To ColN - 2
                If max < Arr01(i + 1) Then
                    max = Arr01(i + 1)
                End If
            Next
        '-----------------------------------
            For i = 0 To ColN - 1
                If Arr01(i) = max Then
                    ii = i
                    Exit For
                End If
            Next
            If x = 1 And max >= 0 Then
                Sheets(PageName02).Cells(VRowN, ColS + ii).Value = "V"
                c_GetRightColInM01 = ColS + ii
            Else
                c_GetRightColInM01 = -1
            End If

End If
End Function

Public Function c_CompareTwoWordInM01(WhichColNum01 As Integer, WhichColNum02 As Integer, U As Integer, L As Integer) As Integer()
'程序功能：
'   对比矩阵01中的两列，如果是8则跳过，如果一致返回1，如果不一致返回0
'程序版本：
'   1.1
'版本修订：
'   1.0 >>> 原始版本,并无颜色标记
    Dim PageName02 As String
    PageName02 = cs_FV("ScreenSheet")
    Dim A As String
    Dim B As String
    Dim C(1 To 2) As Integer
    Dim HowLong As Integer
    Dim i As Integer
    For i = 0 To L - U
        If Sheets(PageName02).Cells(U + i, WhichColNum01).Value <> 8 And Sheets(PageName02).Cells(U + i, WhichColNum02).Value <> 8 Then
            A = A & Trim((Str(Sheets(PageName02).Cells(U + i, WhichColNum01).Value)))
            B = B & Trim((Str(Sheets(PageName02).Cells(U + i, WhichColNum02).Value)))
            HowLong = HowLong + 1
        End If
    Next
    If A = B Then
        C(1) = 1
        C(2) = HowLong
    Else
        C(1) = 0
        C(2) = -1
    End If
    c_CompareTwoWordInM01 = C
End Function

Public Function c_IfRightColInM01ShouldRun(x As Integer)
'程序功能：
'   检查在同一列有没有超过规定的次数，没有返回1，有返回0，同时统计次数，统计次数仅针对上一次也是相同列的情况进行累加，别的都清0
'程序版本：
'   1.1
'版本修订：
'   1.0 >>> 原始版本,并无颜色标记

Dim i As Long
Dim guodu As Long
Dim RowN As Long
Dim ColN As Long
Dim RowS As Long
Dim ColS As Long
Dim PageName02 As String
Dim Row1M0 As Long
Dim Row2M0 As Long
Dim Row3M0 As Long
Dim Row4M0 As Long
Dim M00Fuction1 As String
Dim M00Fuction2 As String
Dim A As Integer
Dim B As Integer
Dim a_1 As Integer
Dim b_1 As Integer
A = 1
B = 1
quanju = 0
a_1 = 1
b_1 = 1
RowN = cs_FV("M01 Rows")
ColN = cs_FV("M01 Cols")
RowS = cs_FV("M01 Row Start Number")
ColS = cs_FV("M01 Col Start Number")
PageName02 = cs_FV("ScreenSheet")

Row1M0 = cs_FV("M00 ROW TOTAL MAX") '判断M00， 单列全局可执行次数，最大，所在行
Row2M0 = cs_FV("M00 ROW TOTAL") '记录M00， 单列全局可执行次数，所在行
Row3M0 = cs_FV("M00 ROW ONE MAX") '判断M00， 单列单次可执行次数，最大，所在行
Row4M0 = cs_FV("M00 ROW ONE") '记录M00， 单列单次可执行次数，所在行
M00Fuction1 = cs_FV("M00 Judge TOTAL MAX") ' 是否启用判断 Total max， Y or N
M00Fuction2 = cs_FV("M00 Judge ONE MAX") ' 是否启用判断 ONE max， Y or N

If x > 0 Then
    If M00Fuction2 = "Y" Then
        If Sheets(PageName02).Cells(Row4M0, x).Value < Sheets(PageName02).Cells(Row3M0, x).Value Then
            a_1 = 1
        Else
            a_1 = 0
        End If
    End If
    If M00Fuction1 = "Y" Then
        If Sheets(PageName02).Cells(Row2M0, x).Value < Sheets(PageName02).Cells(Row1M0, x).Value Then
            b_1 = 1
        Else
            b_1 = 0
        End If
    End If
    If M00Fuction2 = "Y" Then
        If Sheets(PageName02).Cells(Row4M0, x).Value < Sheets(PageName02).Cells(Row3M0, x).Value And a_1 * b_1 = 1 Then
            Sheets(PageName02).Cells(Row2M0, x).Value = Sheets(PageName02).Cells(Row2M0, x).Value + 1
            quanju = 1
            guodu = Sheets(PageName02).Cells(Row4M0, x).Value + 1
            For i = ColS To ColS + ColN - 1
                Sheets(PageName02).Cells(Row4M0, i).Value = 0
            Next
            Sheets(PageName02).Cells(Row4M0, x).Value = guodu
            A = 1
        Else
            A = 0
        End If
    End If
    If M00Fuction1 = "Y" Then
        If Sheets(PageName02).Cells(Row2M0, x).Value < Sheets(PageName02).Cells(Row1M0, x).Value And a_1 * b_1 = 1 Then
            If quanju = 0 Then
                Sheets(PageName02).Cells(Row2M0, x).Value = Sheets(PageName02).Cells(Row2M0, x).Value + 1
            End If
            B = 1
        Else
            B = 0
        End If
    End If
    If A * B = 1 Then
        c_IfRightColInM01ShouldRun = 1
    Else
        c_IfRightColInM01ShouldRun = 0
    End If
ElseIf x < 0 Then
    c_IfRightColInM01ShouldRun = 0
End If
End Function

Public Sub c_IFController()
'程序功能：
'   选择 设置-homepage-接口 页里 judge哪一个，来执行接口程序
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    Dim PageName02 As String
    Dim PageName03 As String
    Dim PageName04 As String
    Dim Range01 As String
    PageName02 = cs_FV("ScreenSheet")
    PageName03 = cs_FV("IFSheet")
    PageName04 = cs_FV("IFInWhichVBAModule")
    Range01 = cs_FV("ShowStatusInWhichCellInScreenSheet")
    T = Timer
    With Sheets(PageName03)
        For i = 2 To .Range("a1000").End(xlUp).Row
            Call cs_ShowJobStatus(Range01, 0)
            If .Range("D" & i).Value = 1 Then
                T2 = Timer
                AA = "'" & PageName04 & "." & .Range("b" & i) & "'"
                Application.Run AA
                DoEvents
                Call cs_Log(AA & " " & Str((Timer - T2)), "Info")
            Else
                Sheets(PageName02).Range(.Range("A" & i).Value).Interior.Color = 6250335
                Sheets(PageName02).Range(.Range("A" & i).Value).Value = 0
                DoEvents
            End If
            Call cs_ShowJobStatus(Range01, 1)
        Next
    End With
    Call cs_Log(" IF Total : " & Str((Timer - T)), "Info")
End Sub
