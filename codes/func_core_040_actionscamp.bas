Attribute VB_Name = "func_core_040_actionscamp"
'-----------------------------------------------
'   All available actions
'   a=action
'-----------------------------------------------
Option Explicit

Public Sub a_X8()
'1�����ܣ�
'1   �ճ���
'1����汾��
'1  1.0
'�汾�޶���
'1   1.0 >>> ԭʼ�汾
End Sub

Public Sub a_AddOne(Whichcount As String)
'1�����ܣ�
'1   Ϊcountԭ�е�ֵ����1
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_WhichCount As String
'2��ֵ
    cv_WhichCount = Whichcount
'2����
    Call cs_AddOne(cv_WhichCount)
End Sub

Public Sub a_ReduceOne(Whichcount As String)
'1�����ܣ�
'1   Ϊcountԭ�е�ֵ����1
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_WhichCount As String
'2��ֵ
    cv_WhichCount = Whichcount
'2����
    Call cs_ReduceOne(cv_WhichCount)
End Sub

Public Sub a_BeZero(Whichcount As String)
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
    Call cs_BeZero(cv_WhichCount)
End Sub

Public Sub a_BeOne(Whichcount As String)
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
    Call cs_BeOne(cv_WhichCount)
End Sub

Public Sub a_ShapeShow(PS As String)
'1���ܣ�
'1   ��ʾ��ע��which����ͼ����
'�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
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
'1���ܣ�
'1   ���ر�ע
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
        On Error Resume Next
        Dim cv_ashape
        For Each cv_ashape In ActiveSheet.Shapes
            If Left(cv_ashape.TextFrame.Characters.Text, 2) = "PS" Then
                cv_ashape.Visible = False
            End If
        Next
        
'2������ڱ�ע̫�ർ�����й���������Ҫ�ݻ�ʹ�����´���
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
'1���ܣ�
'1   ����one�����ܱ���
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
    Call cs_WV("Open Engine", "Y")
    Call cs_Log("One�����Ѿ���", "Info")
End Sub

Public Sub a_CloseOne()
Attribute a_CloseOne.VB_ProcData.VB_Invoke_Func = "S\n14"
'1���ܣ�
'1   ����one���治�ܱ���
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
    Call cs_WV("Open Engine", "N")
    Call cs_Log("One�����Ѿ��ر�", "Info")
End Sub

Public Sub a_OpenActions()
'1���ܣ�
'1   ��ִ�ж�������
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
    Call cs_WV("Actions", "Y")
    Call cs_Log("One����ִ�ж��������Ѿ���", "Info")
End Sub

Public Sub a_CloseActions()
'1���ܣ�
'1   �ر�ִ�ж�������
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
    Call cs_WV("Actions", "N")
    Call cs_Log("One����ִ�ж��������Ѿ��ر�", "Info")
End Sub

Public Sub a_IniExcelName()
'1�����ܣ�
'1   ��excel�����������������ACAK���ĵ�·����д�� setup
'1����汾��
'1   1.3
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'1   1.1 >>> ����excel 2016 365 excel ���ֻ���Щ�ı䣬ԭ���ĳ���ò��������������Ҫ�Ľ����롣
'1   1.2 >>> ����һ���¹��ܣ��Զ���ACAK���ڵ��ļ���λ�ø��µ���core_setup��ҳ�����Excel Path����
'1   1.3 >>> �򻯳���
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
'1�����ܣ�
'1   �Ƿ�򵥻�excelsheetҳ����ʾ ��������ȥ����һ������
'1����汾��
'1   1.1
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'1   1.1 >>> �°汾��������ֵ���ڵ���ҳ��
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
'1�����ܣ�
'1   ��󻯣���С����������excel����
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
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
'1�����ܣ�
'1   ��ʾָ����sheet,ͬʱ����welcomepage��homepage������ҳ����
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
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
            Call a_SimpleExcel   '���´���ʱ�򣬼����Ҫ�󴰿ڵ���ʽ
            Exit For
        End If
    Next
End Sub

Public Sub a_CheckFolder()
'1�����ܣ�
'1   �����setupҳ,ָ�����ļ����Ƿ���ڣ��������򴴽�
'1����汾��
'1   1.1
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'1   1.1 >>> ��Ҫ��̬����setupҳ�г��ֵ�folders ��ȱ����Щ�ļ�����Ҫ����
    Dim cv_excelpath As String
    Dim cv_vname As String
    Dim cv_n As Integer
    Dim cv_i As Integer
    Dim cv_folder01 As String
    Dim cv_folderpath01 As String
    Dim cv_sh
    cv_excelpath = ThisWorkbook.Path
'���ԭ�� setupҳ Ҫ��Ҫ�����ļ���
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
'1�����ܣ�
'1   ��log ҳ��������ݼ��뵽TXT�У�logҳ���
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
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
'1�����ܣ�
'1   ��M00 �е�ֵ��0
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
    Dim cv_i As Long
    Dim cv_ColN As Long
    Dim cv_ColS As Long
    Dim cv_Row2M0 As Long
    Dim cv_Row4M0 As Long
    Dim cv_PageName02 As String
    cv_PageName02 = cs_FV("ScreenSheet")
    cv_ColN = cs_FV("M01 Cols")
    cv_ColS = cs_FV("M01 Col Start Number")
    cv_Row2M0 = cs_FV("M00 ROW TOTAL") '��¼M00�� ����ȫ�ֿ�ִ�д�����������
    cv_Row4M0 = cs_FV("M00 ROW ONE") '��¼M00�� ���е��ο�ִ�д�����������
    For cv_i = cv_ColS To cv_ColS + cv_ColN - 1
        ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row4M0, cv_i).value = 0
    Next
    For cv_i = cv_ColS To cv_ColS + cv_ColN - 1
        ThisWorkbook.Sheets(cv_PageName02).Cells(cv_Row2M0, cv_i).value = 0
    Next
End Sub

Public Sub a_changeOneLoppNumber(loopnumber As Long)
'1�����ܣ�
'1   ����one�������������
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
    Dim cv_loopnumber As Long
    cv_loopnumber = loopnumber
    Call cs_WV("Engine Loop", cv_loopnumber)
    Call cs_Log("����One�����������: " & CStr(cv_loopnumber), "Info")
End Sub

Public Sub a_ExportCode()
'1���ܣ�
'1   ����ACAK�����д���
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
    Call cs_ExportCode
End Sub

Public Sub a_FindPlugin()
'1�����ܣ�
'1   ����Ŀ¼�о���plugin���ܵ�xlsm�ļ�̽���������core_plugin����ʾ��������
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
    Call cs_FindPlugin
End Sub


Public Sub a_IFController()
'1�����ܣ�
'1   ����һ��c_IFController, ��IF sensor ������һ��
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
    Call c_IFController
End Sub

Public Sub a_CheckACAKCore()
'1�����ܣ�
'1   ���ACAKCore�Ƿ�������ҳ���Ƿ񶼴��ڣ������Ƿ񶼴���
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
'2����
Dim yinyong As Variant 'yinyong ��һ��2ά���飬������뵼�������
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
'2��ֵ
cv_i = 1
cv_installed = 0
cv_sheetname = "Core_ACAK_structure"
'2����
'2���ָ��������û�б�����
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
'2���û�е��룬��������
        If cv_installed = 0 Then
            Call cs_Log("Try to load reference: " & .Range("C" & cv_i), "Info")
            Call ThisWorkbook.VBProject.References.AddFromGuid(.Range("C" & cv_i), Val(.Range("D" & cv_i)), Val(.Range("E" & cv_i)))
            .Range("F" & cv_i) = "Y"
        End If
    Next
End With
'2���sheets�Ƿ񶼴���
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
