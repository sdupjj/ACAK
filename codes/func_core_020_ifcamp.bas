Attribute VB_Name = "func_core_020_ifcamp"
'-----------------------------------------------
'   ACAK Interface module
'-----------------------------------------------
Option Explicit
Public Sub if_IFonCo(WhichRange As String, WhichCountTimely As String, WhichCountBenchmark As String)
'1�����ܣ�
'1   eye�ӿڣ����WhichCountTimely >= WhichCountBenchmark,��Ӧ��ֵΪ1����֮��Ϊ0
'1����汾��
'1   1.0
'1�汾�޶���
'1   1.0 >>> ԭʼ�汾
    Dim cv_page02 As String
    Dim cv_CountTimely As Long
    Dim cv_CountBenchmark As Long
    Dim cv_Range01 As String
    cv_Range01 = WhichRange
    cv_page02 = cs_FV("ScreenSheet")
    cv_CountTimely = cs_FindValue(WhichCountTimely)
    cv_CountBenchmark = cs_FindValue(WhichCountBenchmark)
    With ThisWorkbook.Sheets(cv_page02).Range(cv_Range01)
            '2��ʼ��Ϊ��ɫ
            .Interior.Color = RGB(141, 145, 146)
            If cv_CountTimely < cv_CountBenchmark Then
            '2����в�ͬ��ָʾ��Ϊ��ɫ
                .Interior.Color = RGB(207, 1, 37)
                .Font.Color = RGB(255, 255, 255)
                .value = 0
            Else
            '2��ͬΪ��ɫ
                .Interior.Color = RGB(42, 167, 75)
                .Font.Color = RGB(255, 255, 255)
                .value = 1
            End If
    End With
End Sub

