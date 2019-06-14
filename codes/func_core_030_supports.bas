Attribute VB_Name = "func_core_030_supports"
'1-----------------------------------------------
'1   Very basic function / sub in ACAK
'1   cs_ = core support
'1   cv_ = core variables
'1-----------------------------------------------
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Function cs_num2asc2(n As Long) As String
'1�����ܣ�
'1   �������ַ���Ӣ����ĸ
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_n As Long
'2��ֵ
    cv_n = n
'2����
    cs_num2asc2 = Mid(Cells(1, cv_n).Address, 2, IIf(cv_n < 27, 1, 2))
End Function

Public Function cs_FV(Ipt As Variant, Optional WhichSheet As String = "core_setup", Optional WhichCol01 As String = "F", _
                Optional WhichCol02 As String = "G", Optional HowManyRows As Long = 1000) As Variant
'1�����ܣ�
'1   ��ָ���ı��У�ͨ��ָ�����е�ֵ������һ��ָ�����е�ֵ
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_i As Long
    Dim cv_x As Long
    Dim cv_n As Long
    Dim cv_InputValue As Variant
    Dim cv_SheetsName As String
    Dim cv_Col01Name As String
    Dim cv_Col02Name As String
    Dim cv_HowManyRowsInCol01 As Long
'2��ֵ
    cv_InputValue = Ipt
    cv_SheetsName = WhichSheet
    cv_Col01Name = WhichCol01
    cv_Col02Name = WhichCol02
    cv_HowManyRowsInCol01 = HowManyRows
'2����
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
'1�����ܣ�
'1   ��ָ���ı��У�ͨ��ָ�����е�ֵIpt��д����һ��ָ�����е�ֵWpt
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_i As Long
    Dim cv_x As Long
    Dim cv_n As Long
    Dim cv_InputValue As Variant
    Dim cv_WriteValue As Variant
    Dim cv_SheetsName As String
    Dim cv_Col01Name As String
    Dim cv_Col02Name As String
    Dim cv_HowManyRowsInCol01 As Long
'2��ֵ
    cv_InputValue = Ipt
    cv_WriteValue = Wpt
    cv_SheetsName = WhichSheet
    cv_Col01Name = WhichCol01
    cv_Col02Name = WhichCol02
    cv_HowManyRowsInCol01 = HowManyRows
'2����
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
'1�����ܣ�
'1   Ϊcountԭ�е�ֵ����1
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_MiddleCount As Long
    Dim cv_WhichCount As String
'2��ֵ
    cv_WhichCount = Whichcount
'2����
    cv_MiddleCount = cs_FV(cv_WhichCount, cs_FV("CountSheet"), "A", "B")
    Call cs_WV(cv_WhichCount, cv_MiddleCount + 1, cs_FV("CountSheet"), "A", "B")
End Sub

Public Sub cs_ReduceOne(Whichcount As String)
'1�����ܣ�
'1   Ϊcountԭ�е�ֵ����1
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_MiddleCount As Long
    Dim cv_WhichCount As String
'2��ֵ
    cv_WhichCount = Whichcount
'2����
    cv_MiddleCount = cs_FV(cv_WhichCount, cs_FV("CountSheet"), "A", "B")
    Call cs_WV(cv_WhichCount, cv_MiddleCount - 1, cs_FV("CountSheet"), "A", "B")
End Sub

Public Sub cs_BeZero(Whichcount As String)
'1�����ܣ�
'1   ����count��ֵΪ0
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_WhichCount As String
'2��ֵ
    cv_WhichCount = Whichcount
'2����
    Call cs_WV(cv_WhichCount, 0, cs_FV("CountSheet"), "A", "B")
End Sub

Public Sub cs_BeOne(Whichcount As String)
'1�����ܣ�
'1   ����count��ֵΪ1
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_WhichCount As String
'2��ֵ
    cv_WhichCount = Whichcount
'2����
    Call cs_WV(cv_WhichCount, 1, cs_FV("CountSheet"), "A", "B")
End Sub

Public Sub cs_BeValue(Whichcount As String, value As Variant)
'1�����ܣ�
'1   д������ҳ��� whichcount �Աߵ���ֵ���value
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_WhichCount As String
    Dim cv_Value As Variant
'2��ֵ
    cv_WhichCount = Whichcount
    cv_Value = value
'2����
    Call cs_WV(cv_WhichCount, cv_Value, cs_FV("CountSheet"), "A", "B")
End Sub

Public Function cs_FindValue(Whichcount As String) As Variant
'1�����ܣ�
'1   ���ֵ�whichcount �Աߵ���ֵ
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_WhichCount As String
'2��ֵ
    cv_WhichCount = Whichcount
'2����
    cs_FindValue = cs_FV(cv_WhichCount, cs_FV("CountSheet"), "A", "B")
End Function

Public Sub cs_TakeAction(whichpagename As String, whichcol As String, Optional Rows As Long = 2, Optional RowE As Long = 999, _
                                            Optional ColorCellsOrNot As String = "N", Optional ShowErrorOrNot As String = "N")
'1�����ܣ�
'1   ��ĳҳ��ĳ�ж��� ����ִ��
'1����汾��
'1   1.1
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾,������ɫ���
'1   1.1 >>> ��״̬����ʾ����
'2����
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
    
'2��ֵ

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
'2����
'2�õ����õ�����action
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
            '2 �õ���ǰҪִ�е�action����
            If InStr(.value, " ") > 0 Then
                cv_C = Left(.value, InStr(.value, " ") - 1)
            Else
                cv_C = .value
            End If
            '2�鿴�����Ҫִ�е�action�Ƿ���ģ���п��ñ����֣����������������־��д��error
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
            Call cs_Log(.value & " ��ʱ�� " & Str(Application.Text((Timer - cv_T01), "0.0000")), "Debug")
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
'1�����ܣ�
'1   ����һ�δ���
'1����汾��
'1   1.1
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
Dim cv_AA As String
'2��ֵ
cv_AA = AA
'2����
On Error GoTo debugprint
    Application.Run cv_AA
    Call cs_Log("Run" & cv_AA, "Info")
    Exit Sub
debugprint:
    Call cs_Log("Run" & cv_AA & " works error.", "Error")
End Sub

Public Sub cs_Log(sts As Variant, Optional Level_Error_Warning_Debug_Info_Print As String = "Print")
'1�����ܣ�
'1   ��¼sts��core_log��
'1����汾��
'1   1.0
'1   1.1
'1   1.2
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'1   1.1 >>> ����һ������ ȷ�� ����ʲô����¼�¼log
'1    1.2>>> ��ԭ����Ҫ��ɢ���� ��־��¼����Level_Error_Warning_Debug_Info_Print ���ڼ������������������
'2����
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
'2��ֵ
    cv_PageName = cs_FV("LogSheet")
    cv_Record = sts
    cv_LogLevel = Level_Error_Warning_Debug_Info_Print
    cv_array01() = Split(cs_FV("LogLevel"), ";")
    cv_i = UBound(cv_array01)
    cv_array02() = Split(cs_FV("ShowInInfoWindowsLogLevel"), ";")
    cv_ii = UBound(cv_array02)
'2����
'2��core_log ҳ������ʾ
    For cv_n = 0 To cv_i
        If cv_array01(cv_n) = cv_LogLevel Then
            '2��¼log��¼���ڼ���
            '2 count03  ���ֵӦ����ACAK������ر�ʱ���������ó� 0 ������
            Call cs_AddOne("count03")
            cv_A = cs_FindValue("count03") + 1
            '2����log��¼
            cv_B = Now() & " <" & cv_LogLevel & "> " & cv_Record
            '2���浽log��¼ҳ��
            ThisWorkbook.Sheets(cv_PageName).Range("B" & cv_A).value = cv_B
            Exit For
        End If
    Next
'2��info����������ʾ
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
'1�����ܣ�
'1   ����ʾ��ҳָ����Ԫ����ʾ��ɫ/Done�� ��ɫ/Working
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_PageName02 As String
    Dim cv_ShowStatusInWhichCell As String
    Dim cv_w0d1 As Integer
'2��ֵ
    cv_PageName02 = cs_FV("ScreenSheet")
    cv_ShowStatusInWhichCell = ShowStatusInWhichCell
    cv_w0d1 = w0d1
'2����
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
'1�����ܣ�
'1   ���ݸ����Ĵ�������wName�õ���Ӧ�Ĵ�������
'1����汾��
'1  1.1
'1�汾�޶���
'1  1.0 >>> ԭʼ�汾
'1  1.1 >>> ɾ�����ô���
'2����
    Dim cv_wName As String
'2��ֵ
    cv_wName = wName
'2����
    On Error GoTo error1
    cs_FindWnd = FindWindow(vbNullString, cv_wName)
    Call cs_Log(Now() & " FindWnd��" & "���壺 " & cv_wName & " ���:  " & cs_FindWnd, "Debug")
    Exit Function
error1:
    Call cs_Log(Now() & " ��ͼ���� ���壺" & cv_wName & " �ľ�� ��FindWnd��������", "Error")
End Function

Public Sub cs_FindPlugin()
'1�����ܣ�
'1   ����Ŀ¼�о���plugin���ܵ�xlsm�ļ�̽���������core_plugin����ʾ
'1����汾��
'1  1.0
'1�汾�޶���
'1  1.0 >>> ԭʼ�汾
'2����
    Dim cv_f As String
    Dim cv_p As String
    Dim cv_d
    Dim cv_ar
    Dim cv_i As Integer
'2��ֵ
    Set cv_d = CreateObject("Scripting.Dictionary")
    cv_p = ThisWorkbook.Path & cs_FV("Folder For Plugin")
    cv_f = Dir(cv_p & "Plugin*.xlsm")
'2����
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
'1�����ܣ�
'1   ������к�������
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_ExistingValue As Long
'2��ֵ
'2����
    cv_ExistingValue = cs_FV("Plugin No. :", cs_FV("Sheet For Plugin"), "F", "G")
    Call cs_WV("Plugin No. :", (cv_ExistingValue + 1), cs_FV("Sheet For Plugin"), "F", "G")
    cs_GenPluginNumber = cv_ExistingValue + 1
End Function

Public Sub cs_LoadPlugin()
'1�����ܣ�
'1   ����Plugin
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾

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
'2ȷ���Ѿ����صĲ�����ᱻ�ٴμ���\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        For cv_ii = 5 To 100
            If ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I" & cv_ii).value = cv_PluginName Then
                MsgBox "Seems Plugin: " & cv_PluginName & " is loaded. Please check it at first."
                GoTo out
            End If
        Next
 '2////////////////////////////////////////////////////////////////////////////
        cv_PluginAddress = cs_FV("Excel Path") & cs_FV("Folder For Plugin") & cv_PluginName
        cv_PluginNumber = cs_GenPluginNumber()
'2�������п�ͷΪs_uc��sheet������vba���룩\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        Set cv_PluginXLSM = Workbooks.Open(cv_PluginAddress)
        For Each cv_sh In cv_PluginXLSM.Sheets
            If Left(cv_sh.name, 4) = "s_uc" Then
                cv_PluginXLSM.Sheets(cv_sh.name).Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(cv_sh.name).name = "PI" & cv_PluginNumber & "_" & cv_sh.name
'2��core_homepage������ÿ��sheet�İ�ť
                Call cs_CreateShape("PI" & cv_PluginNumber & "_" & cv_sh.name)
            End If
        Next
'2////////////////////////////////////////////////////////////////////////////
'2��������vbaģ��\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'2����ģ�� ǰ׺���� ��PI�����_��
        For Each cv_ovbcomp In cv_PluginXLSM.VBProject.VBComponents
            Select Case cv_ovbcomp.type
                Case vbext_ct_StdModule
                    Set cv_ovbcompUC = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_StdModule)
                    cv_ovbcompUC.name = "PI" & cv_PluginNumber & "_" & cv_ovbcomp.name
'2�������ģ���д���
                    cv_sfunc = ""
                    With cv_PluginXLSM.VBProject.VBComponents(cv_ovbcomp.name).CodeModule
                        For cv_lindex = 1 To .CountOfLines
                            cv_A = .Lines(cv_lindex, 1)
'2�������ģ��Ĵ������ᵽ�� s_uc_ and f_uc_ ����ǰ׺ PIXXX_
                            cv_MingganZi01 = "s_uc_"
                            cv_MingganZi02 = "PI" & cv_PluginNumber & "_" & "s_uc_"
                            cv_MingganZi03 = "f_uc_"
                            cv_MingganZi04 = "PI" & cv_PluginNumber & "_" & "f_uc_"
                            cv_A = Replace(cv_A, cv_MingganZi01, cv_MingganZi02)
                            cv_A = Replace(cv_A, cv_MingganZi03, cv_MingganZi04)
                            cv_sfunc = cv_sfunc & cv_A & Chr(10)
                        Next
                    End With
'2д����뵽ACAK Core ģ����
                    With cv_ovbcompUC.CodeModule
                        cv_lLinestart = .CountOfLines + 1
                        .InsertLines cv_lLinestart, cv_sfunc
                    End With
'2������е�if �� action ����f_core_020_ifcamp ��f_040_actionscamp
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
'2ɾ�����module
        For Each cv_ovbcomp In ThisWorkbook.VBProject.VBComponents
            Select Case cv_ovbcomp.type
                Case vbext_ct_StdModule 'case 1
                    If Left(cv_ovbcomp.name, 13 + cv_length) = "PI" & CStr(cv_PluginNumber) & "_func_debug" Then
                        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(cv_ovbcomp.name)
                    End If
            End Select
        Next

'2�رղ��

        cv_PluginXLSM.Save
        cv_PluginXLSM.Close
'2ע������ACAK
        cv_NextLoadPluginRow = ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I1000").End(xlUp).Row + 1
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I" & cv_NextLoadPluginRow).value = cv_PluginName
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("J" & cv_NextLoadPluginRow).value = cv_PluginNumber
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("K" & cv_NextLoadPluginRow).value = ""
'2ִ��s_uc_plugin_setupҳ�С�After Loaded������
''2�ȰѴ���PIXXX��
'        With Sheets("PI" & CStr(cv_PluginNumber) & "_s_uc_plugin_setup")
'            For cv_iii = 2 To .Range("D1000").End(xlUp).Row
'                If .Range("D" & cv_iii).value <> "" Then
'                    .Range("D" & cv_iii).value = "PI" & CStr(cv_PluginNumber) & "_" & .Range("D" & cv_iii).value
'                End If
'            Next
'        End With
'2ִ��
'Call cs_TakeAction("PI" & CStr(cv_PluginNumber) & "_s_uc_plugin_setup", "D")
    End If
Next
'2ˢ����action��IF list,�½�����ļ���
Call cs_FindActions
Call cs_FindIFs
Call a_CheckFolder
Call a_ShowPage("core_plugin")

MsgBox "Done"
out:
Call a_ShowPage("core_plugin")
End Sub

Public Sub cs_unLoadPlugin()
'�����ܣ�
'   �Ƴ�Plugin
'����汾��
'   1.0
'�汾�޶���
'   1.0 >>> ԭʼ�汾
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
'2ɾ��ť
        Call cs_RemoveShape("PI" & cv_PluginNumber)
'2ɾsheet
        cv_length = Len(CStr(cv_PluginNumber))
        For Each cv_sh In ThisWorkbook.Sheets
            If Left(cv_sh.name, 2 + cv_length) = "PI" & cv_PluginNumber Then
                Application.DisplayAlerts = False
                ThisWorkbook.Sheets(cv_sh.name).Delete
                Application.DisplayAlerts = True
            End If
        Next
'2ɾ�����module
        For Each cv_ovbcomp In ThisWorkbook.VBProject.VBComponents
            Select Case cv_ovbcomp.type
                Case vbext_ct_StdModule 'case 1
                    If Left(cv_ovbcomp.name, 2 + cv_length) = "PI" & cv_PluginNumber Then
                        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(cv_ovbcomp.name)
                    End If
            End Select
        Next
'2ɾif �� action�����ӹ����Ĵ���
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
'2ɾ��ť
        Call cs_RemoveShape("PI" & cv_PluginNumber)
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("I" & cv_ii).value = ""
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("J" & cv_ii).value = ""
        ThisWorkbook.Sheets(cv_SheetForPlugin).Range("K" & cv_ii).value = ""
    End If
Next

'ˢ����action and IF list
Call cs_FindActions
Call cs_FindIFs
Call a_ShowPage("core_plugin")
MsgBox "Done. *** The used actions in M02 and used if in core_if should be removed manually. "
End Sub

Public Function cs_ucSubFuncToACAK(ucSubFuncSentence As String, PluginNumber As Long, Remark As String, PIXXXModuleName As String)
'1�����ܣ�
'1   ������е�Ŀ����ת��ΪACAK�������õ�sub �� function, Ŀǰֻ������SUB
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
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
    '�õ�array01 �ж����� ubound +1
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
    cv_Sentence02 = "'������:" & Chr(10)
    cv_Sentence03 = cv_Remark & Chr(10)
    cv_Sentence04 = "    Call " & "PI" & cv_PluginNumber & "_" & cv_PIXXXModuleName & "." & cv_SubFuncName & "(" & cv_NewArguments & ")" & Chr(10)
    cv_Sentence05 = "End Sub" & Chr(10)
    cs_ucSubFuncToACAK = cv_Sentence01 & cv_Sentence02 & cv_Sentence03 & cv_Sentence04 & cv_Sentence05
End Function

Public Sub cs_FindActions()
'1�����ܣ�
'1   ��"func_core_040_actionscamp"������action����ʾ�ڡ�core_actions��
'1����汾��
'1   1.0
'1�汾�޶���
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
'1�����ܣ�
'1   ��"func_core_020_ifcamp"������action����ʾ�ڡ�core_ifcamp��
'1����汾��
'1   1.0
'1�汾�޶���
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
'1�����ܣ�
'1   ��core_homepage����ʾ �����ز����sheet
'1����汾��
'1   1.0
'1�汾�޶���
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
'1�����ܣ�
'1   ��core_homepage��ɾ��nameΪshapename�İ�ť
'1����汾��
'1   1.0
'1�汾�޶���
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
'1�����ܣ�
'1   ��ʾErrorForm
'1����汾��
'1   1.0
'1�汾�޶���
'1    1.0
    Info.Show 0
End Sub

