Attribute VB_Name = "func_core_990_debug"
'-----------------------------------------------
'   ACAK debug codes
'   cs=core support
'-----------------------------------------------
Option Explicit
Public Sub cs_ExportCode()
'1���ܣ�
'1   ����ACAK�����д���
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
    Dim cv_strExportFolderPath As String
    Dim cv_ovbcomp As VBComponent
    Dim cv_oExportFolder As Object
    Dim cv_strExtName As String
    Dim cv_strVBCompName As String
    Dim cv_ovbproj
    Dim cv_zd
    Dim cv_sht
    Dim cv_zdkeys
    Dim cv_i
    
    cv_strExportFolderPath = ThisWorkbook.Path & cs_FV("Folder For Codes")
    
    Set cv_zd = CreateObject("scripting.dictionary")
    
    For Each cv_ovbcomp In ThisWorkbook.VBProject.VBComponents
        If Not cv_zd.exists(cv_ovbcomp.name) Then
            Select Case cv_ovbcomp.type
                Case vbext_ct_StdModule 'case 1
                    cv_zd(cv_ovbcomp.name) = cv_ovbcomp.name & ".bas"
                Case vbext_ct_ClassModule 'case 2
                    cv_zd(cv_ovbcomp.name) = cv_ovbcomp.name & ".cls"
                Case vbext_ct_MSForm 'case 3
                    cv_zd(cv_ovbcomp.name) = cv_ovbcomp.name & ".frm"
                Case vbext_ct_Document 'case 100
                    If cv_ovbcomp.name <> "ThisWorkbook" Then
                        For Each cv_sht In ThisWorkbook.Worksheets
                            If cv_sht.CodeName = cv_ovbcomp.name Then
                                cv_zd(cv_ovbcomp.name) = cv_sht.name & ".txt"
                            End If
                        Next
                    Else
                        cv_zd(cv_ovbcomp.name) = cv_ovbcomp.name & ".txt"
                    End If
            End Select
        End If
    Next
    
    Set cv_ovbproj = ThisWorkbook.VBProject
    cv_zdkeys = cv_zd.keys
    For cv_i = 0 To cv_zd.Count - 1
        cv_strVBCompName = cv_zdkeys(cv_i)
        cv_ovbproj.VBComponents(cv_strVBCompName).Export cv_strExportFolderPath & "\" & cv_zd(cv_strVBCompName)
Next
MsgBox "ACAK Codes Exported!"
End Sub

Public Sub cs_OpenFatherFolder()
'1���ܣ�
'1   ��ACAK���ڵ��ļ���
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
    Dim cv_Path As String
    cv_Path = ThisWorkbook.Path
    Shell "explorer.exe " & cv_Path, vbNormalFocus
End Sub

Public Sub cs_ExportReference()
'1���ܣ�
'1   ����ACAK����������
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
Dim cv_i As Long
Dim cv_refed
Dim cv_sheetname As String
cv_i = 4
cv_sheetname = "Core_ACAK_structure"

With Sheets(cv_sheetname)
    .Range("B4:F100").ClearContents
    For Each cv_refed In ThisWorkbook.VBProject.References
        .Cells(cv_i, 2) = cv_refed.name
        .Cells(cv_i, 3) = cv_refed.Guid
        .Cells(cv_i, 4) = cv_refed.Major
        .Cells(cv_i, 5) = cv_refed.Minor
        cv_i = cv_i + 1
    Next
End With

End Sub


Public Sub cs_ExportSettings()
'1���ܣ�
'1   ����ACAK������ҳ�����õ��������ļ���\Settings backup\
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
'2����
    Dim cv_xmlName As String 'XML����
    Dim cv_xmlDoc As DOMDocument '�����ܱ�
    Dim cv_xmlPI As IXMLDOMProcessingInstruction '�����ܱ����Ի����Щ����ֵ
    Dim cv_xmlRoot As IXMLDOMElement '�ܱ�����
    Dim cv_xmlVoucher As IXMLDOMElement 'ÿ������
    Dim cv_R As Long
    Dim cv_C As Long
    Dim cv_MergeCellsArea As String
    Dim cv_sh As Worksheet
    Dim cv_Rows As Long
    Dim cv_Columns As Long
    Dim cv_Path As String
    Dim cv_XMLTempPath As String
    Dim cv_SettingsPath As String
    Dim cv_SettingZIPFileName As String
    Dim cv_AddressA As Variant
    Dim cv_AddressB As Variant
    
    cv_SettingZIPFileName = Application.Text(Now(), "yyyymmddhhmmss") & ".zip"
    cv_SettingsPath = ActiveWorkbook.Path & cs_FV("Folder For Settings")
    cv_XMLTempPath = ActiveWorkbook.Path & cs_FV("Folder For Settings") & "Tempp"
    cv_AddressA = Trim(cv_XMLTempPath)
    cv_AddressB = Trim(cv_SettingsPath & cv_SettingZIPFileName)
    
    '2ɾ����ʱ�ļ���
    If Dir(cv_XMLTempPath, vbDirectory) <> "" Then
        Kill cv_XMLTempPath & "\*.*"
        RmDir cv_XMLTempPath
    End If
    '2������ʱ�ļ���
    MkDir cv_XMLTempPath
    For Each cv_sh In ThisWorkbook.Worksheets
        If cv_sh.Range("a1").value <> "" Then
            cv_sh.Range("a1").value = 1
        End If
    Next
    
    For Each cv_sh In ThisWorkbook.Worksheets
        With cv_sh
            Set cv_xmlDoc = New DOMDocument
            Set cv_xmlPI = cv_xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""utf-8""")
            cv_xmlDoc.appendChild cv_xmlPI
            cv_xmlName = .name
            Set cv_xmlRoot = cv_xmlDoc.createElement(.name) '<<<<<<<<<<<<<<<<<<<<<<<<<<<���
            cv_xmlDoc.appendChild cv_xmlRoot
            cv_Rows = .UsedRange.Rows.Count
            cv_Columns = .UsedRange.Columns.Count
            Set cv_xmlVoucher = cs_AddChild(cv_xmlDoc, cv_xmlRoot, "CellValuesR" & cv_Rows & "C" & cv_Columns)    '<<<<<<<<<<<<<<<<<<<<<<<<<<<ÿ����Ԫ���ֵ
            For cv_R = 1 To cv_Rows
                For cv_C = 1 To cv_Columns
                    cs_AddEntry cv_xmlDoc, cv_xmlVoucher, "CV", IIf(.Cells(cv_R, cv_C).Formula <> "", .Cells(cv_R, cv_C).Formula, .Cells(cv_R, cv_C).value)
                Next
            Next
            
            Set cv_xmlVoucher = cs_AddChild(cv_xmlDoc, cv_xmlRoot, "CellMerge")     '<<<<<<<<<<<<<<<<<<<<<<<<<<<��Ԫ��ϲ����
            For cv_R = 1 To cv_Rows
                For cv_C = 1 To cv_Columns
                    If .Cells(cv_R, cv_C).MergeCells Then
                        cv_MergeCellsArea = .Cells(cv_R, cv_C).MergeArea.Address
                        cs_AddEntry cv_xmlDoc, cv_xmlVoucher, "Area", cv_MergeCellsArea
                    End If
                Next
            Next
            Set cv_xmlVoucher = cs_AddChild(cv_xmlDoc, cv_xmlRoot, "CellColor")     '<<<<<<<<<<<<<<<<<<<<<<<<<<<��Ԫ����ɫ
            For cv_R = 1 To cv_Rows
                For cv_C = 1 To cv_Columns
                        cs_AddEntry cv_xmlDoc, cv_xmlVoucher, "CC", .Cells(cv_R, cv_C).Interior.Color
                Next
            Next
            Set cv_xmlVoucher = cs_AddChild(cv_xmlDoc, cv_xmlRoot, "CellFontNameColorBoldSize")     '<<<<<<<<<<<<<<<<<<<<<<<<<<<��Ԫ������;��ɫ;��ϸ;��С
            For cv_R = 1 To cv_Rows
                For cv_C = 1 To cv_Columns
                        cs_AddEntry cv_xmlDoc, cv_xmlVoucher, "FN", .Cells(cv_R, cv_C).Font.name & "_ " & .Cells(cv_R, cv_C).Font.Color & "_ " & .Cells(cv_R, cv_C).Font.Bold & "_" & .Cells(cv_R, cv_C).Font.Size
                Next
            Next
            Set cv_xmlVoucher = cs_AddChild(cv_xmlDoc, cv_xmlRoot, "CellHorizontalAlignmentVerticalAlignment")     '<<<<<<<<<<<<<<<<<<<<<<<<<<<��Ԫ��ˮƽ���о������;��ֱ���о��Ͼ���
            For cv_R = 1 To cv_Rows
                For cv_C = 1 To cv_Columns
                        cs_AddEntry cv_xmlDoc, cv_xmlVoucher, "HA", .Cells(cv_R, cv_C).HorizontalAlignment & "_" & .Cells(cv_R, cv_C).VerticalAlignment
                Next
            Next
            Set cv_xmlVoucher = cs_AddChild(cv_xmlDoc, cv_xmlRoot, "CellBorderLineStyle")     '<<<<<<<<<<<<<<<<<<<<<<<<<<<��Ԫ���������µı߿���ʽ
            For cv_R = 1 To cv_Rows
                For cv_C = 1 To cv_Columns
                        cs_AddEntry cv_xmlDoc, cv_xmlVoucher, "BLS", .Cells(cv_R, cv_C).Borders(xlEdgeLeft).LineStyle & "_" & .Cells(cv_R, cv_C).Borders(xlEdgeRight).LineStyle & "_" & .Cells(cv_R, cv_C).Borders(xlEdgeTop).LineStyle & "_" & .Cells(cv_R, cv_C).Borders(xlEdgeBottom).LineStyle
                Next
            Next
            Set cv_xmlVoucher = cs_AddChild(cv_xmlDoc, cv_xmlRoot, "CellBorderColor")     '<<<<<<<<<<<<<<<<<<<<<<<<<<<��Ԫ���������µı߿���ɫ
            For cv_R = 1 To cv_Rows
                For cv_C = 1 To cv_Columns
                        cs_AddEntry cv_xmlDoc, cv_xmlVoucher, "BC", .Cells(cv_R, cv_C).Borders(xlEdgeLeft).Color & "_" & .Cells(cv_R, cv_C).Borders(xlEdgeRight).Color & "_" & .Cells(cv_R, cv_C).Borders(xlEdgeTop).Color & "_" & .Cells(cv_R, cv_C).Borders(xlEdgeBottom).Color
                Next
            Next
            Set cv_xmlVoucher = cs_AddChild(cv_xmlDoc, cv_xmlRoot, "RowHeightColWidth")     '<<<<<<<<<<<<<<<<<<<<<<<<<<<�����иߺ��и�
            For cv_R = 1 To cv_Rows
                For cv_C = 1 To cv_Columns
                        cs_AddEntry cv_xmlDoc, cv_xmlVoucher, "RHCW", .Cells(cv_R, cv_C).Height & "_" & .Cells(cv_R, cv_C).Width
                Next
            Next
            cv_Path = cv_XMLTempPath & "\" & .name & ".xml"
            cv_xmlDoc.Save cv_Path
        End With
   Next
   
'    MsgBox "ACAK Core Sheets Exported!"
    Call cs_CreateZipFile(cv_AddressA, cv_AddressB)
    '2ɾ����ʱ�ļ���
    If Dir(cv_XMLTempPath, vbDirectory) <> "" Then
        Kill cv_XMLTempPath & "\*.*"
        RmDir cv_XMLTempPath
    End If
    MsgBox "ACAK Settings Exported!"
End Sub
Public Function cs_AddChild(ByVal xmlDoc As DOMDocument, _
                  ByVal xmlParent As IXMLDOMElement, _
                  ByVal tagName As String _
                 ) As IXMLDOMElement
    Dim cv_xmlChild As IXMLDOMElement

    Set cv_xmlChild = xmlDoc.createElement(tagName)
    xmlParent.appendChild cv_xmlChild

    Set cs_AddChild = cv_xmlChild
End Function
Function cs_AddTextChild(ByVal xmlDoc As DOMDocument, _
                      ByVal xmlParent As IXMLDOMElement, _
                      ByVal tagName As String, _
                      ByVal Text As String _
                     ) As IXMLDOMElement
    Dim cs_xmlChild As IXMLDOMElement
    Dim cs_xmlText  As IXMLDOMText

    Set cs_xmlChild = xmlDoc.createElement(tagName)
    xmlParent.appendChild cs_xmlChild
    cs_xmlChild.Text = Text

    Set cs_AddTextChild = cs_xmlChild
End Function
Sub cs_AddEntry(ByVal xmlDoc As DOMDocument, _
             ByVal xmlParent As IXMLDOMElement, _
             ByVal name As String, _
             ByVal value As String)
    cs_AddTextChild xmlDoc, xmlParent, name, value
End Sub

Public Sub cs_CreateZipFile(folderToZipPath As Variant, zippedFileFullName As Variant)
'1���ܣ�
'1  ��ָ���ļ��д����ָ�����ֵ�zip�ļ�
'1�汾��
'1   1.0
'1ÿһ���޶���
'1   1.0 >>> ԭʼ�汾
'2����
Dim ShellApp As Object

'Create an empty zip file
Open zippedFileFullName For Output As #1
Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
Close #1

'Copy the files & folders into the zip file
Set ShellApp = CreateObject("Shell.Application")
ShellApp.Namespace(zippedFileFullName).CopyHere ShellApp.Namespace(folderToZipPath).items

'Zipping the files may take a while, create loop to pause the macro until zipping has finished.
On Error Resume Next
Do Until ShellApp.Namespace(zippedFileFullName).items.Count = ShellApp.Namespace(folderToZipPath).items.Count
    Application.Wait (Now + TimeValue("0:00:01"))
Loop
On Error GoTo 0

End Sub

