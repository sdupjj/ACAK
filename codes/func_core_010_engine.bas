Attribute VB_Name = "func_core_010_engine"
'-----------------------------------------------
'   ACAK One Engine module
'-----------------------------------------------
Option Explicit
Public Sub c_One()
'1�����ܣ�
'1   ������������
'1   �Ƿ������� ��0��һ���ܱ����У�֮�� one���濪�عر�>��������ת�����Ѿ��ﵽ=δ�ҵ�����ִ�е���>���д����Ѿ��ﵽ
'1����汾��
'1   1.1
'1�汾�޶���
'1   1.1 >>> ���Լ�����
'1   1.0 >>> ԭʼ�汾
'2����
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
'2��ֵ
    cv_Vrow = cs_FV("VinScreenRow")
    cv_PageName02 = cs_FV("ScreenSheet")
    cv_PageName03 = cs_FV("EngineSetupSheet")
    cv_ActionsCamp = cs_FV("ActionsInWhichVBAModule")
    cv_Rows = cs_FV("M02 Row Start Number")
    cv_RowSS = cs_FV("M02 Rows") '�õ��ж���ָ����Ҫ����
    cv_EngLoopNumber = cs_FV("Engine Loop")
    cv_CellDoneWorking = cs_FV("ShowStatusInWhichCellInScreenSheet")
'2����
    Call cs_ShowJobStatus(cv_CellDoneWorking, 0)
    For cv_n = 0 To cv_EngLoopNumber + 1
        Select Case cv_n
            '2 One��0�����У���ʼ������
            Case 0
                Call cs_TakeAction(cv_PageName03, "D")
             '2 One ��1�����涨��������
            Case 1 To cv_EngLoopNumber
                '2 ��������Ƿ������
                cv_EngYN = cs_FV("Open Engine")
                '2 ��������ܷ�����ִ�ж���
                cv_BehYN = cs_FV("Actions")
                If cv_EngYN = "Y" Then
                    '2 ���� ACAK�򿪺� One�������д���
                    Call cs_AddOne("count01")
                    '2 ���� One���·��������д���
                    Call cs_AddOne("count02")
                    cv_T07 = Timer
                    '2 ����ÿ��IF<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
                    Call c_IFController
                    '2 "V"��ʶ ���������
                    ThisWorkbook.Sheets(cv_PageName02).Range("G" & cv_Vrow & ":AG" & cv_Vrow).ClearContents
                    '2 �ҳ���V��Ӧ������һ�У�������������кţ����֣�
                    cv_x = c_GetRightColInM01()
                    DoEvents
                    '2 �����ṩ���кţ��ж��Ƿ��Ѿ�����������д��� �����������򷵻� 1�� �����������򷵻�0
                    cv_y = c_IfRightColInM01ShouldRun(cv_x)
                    If cv_x > 0 And cv_y > 0 And cv_BehYN = "Y" Then '��������
                            Call cs_TakeAction(cv_PageName02, cs_num2asc2(cv_x), cv_Rows, cv_Rows + cv_RowSS - 1, "Y")
                            Call cs_Log("One Engine Loop ��ʱ�� " & Str(Application.Text((Timer - cv_T07), "0.0000")), "Print")
                    ElseIf cv_x > 0 And cv_y > 0 And cv_BehYN = "N" Then 'One����������Ϊ������ִ��
                        Call cs_Log("ACAK������Ϊ������������ִ�С�", "Info")
                    ElseIf cv_x < 0 Then '�Ҳ�����Ӧ����'
                        Call cs_Log("M01���Ҳ�����Ӧ���У�����������ִ�С�", "Info")
                        Exit For
                    ElseIf cv_y = 0 Then
                        Call cs_Log("M01 & M02 ��Ӧ���г����˹涨ѭ������������������ִ�С�", "Info")
                        Exit For
                    End If
                Else 'One���汻����δ���ɱ�ִ��
                    Call cs_Log("One�����Ѿ����ر�", "Info")
                    Exit For
                End If
            Case cv_EngLoopNumber + 1
            '2 One�Ѿ�����run��ϣ���Ҫ��β
                Call cs_TakeAction(cv_PageName03, "H")
                Call cs_Log("One�����Ѿ��ﵽ���õ����ѭ���������������С�", "Info")
        End Select
    Next
    Call cs_ShowJobStatus(cv_CellDoneWorking, 1)
End Sub

Public Function c_GetRightColInM01() As Long
'1�����ܣ�
'1   ͨ���Ա����Ľ������M01���ҵ���Ӧ����,���ڶ�Ӧ����ͷ�ϱ����V"
'1����汾��
'1   1.1
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾,������ɫ���
'1   1.1 >>> debug
'����
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
'��ֵ
cv_IFColN = cs_FV("IFdisplayinScreenCol")
cv_VRowN = cs_FV("VinScreenRow")
cv_RowN = cs_FV("M01 Rows")
cv_ColN = cs_FV("M01 Cols")
cv_Rows = cs_FV("M01 Row Start Number")
cv_ColS = cs_FV("M01 Col Start Number")
cv_JudgeWay = cs_FV("EngineColumnSelectMethod")
cv_PageName02 = cs_FV("ScreenSheet")
'����

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
        'ȷ���з�����������
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
Call cs_Log("Found Right Column Total " & Str(Application.Text((Timer - cv_T03), "0.0000")), "Debug") '�鿴�ҵ���ȷ�е�ʱ��
End Function

Public Function c_CompareTwoWordInM01(WhichColNum01 As Long, WhichColNum02 As Long, U As Long, L As Long) As Long()
'1�����ܣ�
'1   �ԱȾ���01�е����У������8�����������һ�·���1�������һ�·���0
'1����汾��
'1   1.1
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾,������ɫ���
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
'1�����ܣ�
'1   ��� x�� ��û�г����涨�Ĵ�����û�з���1���з���0��ͬʱͳ�ƴ�����ͳ�ƴ����������һ��Ҳ����ͬ�е���������ۼӣ���Ķ���0
'1����汾��
'1   1.1
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾,������ɫ���
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

cv_Row1M0 = cs_FV("M00 ROW TOTAL MAX") '�ж�M00�� ����ȫ�ֿ�ִ�д��������������
cv_Row2M0 = cs_FV("M00 ROW TOTAL") '��¼M00�� ����ȫ�ֿ�ִ�д�����������
cv_Row3M0 = cs_FV("M00 ROW ONE MAX") '�ж�M00�� ���е��ο�ִ�д��������������
cv_Row4M0 = cs_FV("M00 ROW ONE") '��¼M00�� ���е��ο�ִ�д�����������
cv_M00Fuction1 = cs_FV("M00 Judge TOTAL MAX") ' �Ƿ������ж� Total max�� Y or N
cv_M00Fuction2 = cs_FV("M00 Judge ONE MAX") ' �Ƿ������ж� ONE max�� Y or N

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
'1�����ܣ�
'1   ѡ�� core_if ҳ�� judge��һ�� �ж��У���ִ����ؽӿڳ�����Ϊ1�������
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
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
'1�����ܣ�
'1   ����core_screen �е�checkstatus��ť
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
    Dim cv_CellDoneWorking As String
    
    cv_CellDoneWorking = cs_FV("ShowStatusInWhichCellInScreenSheet")
    
    Call cs_ShowJobStatus(cv_CellDoneWorking, 0)
    Call c_IFController
    Call cs_ShowJobStatus(cv_CellDoneWorking, 1)
End Sub

