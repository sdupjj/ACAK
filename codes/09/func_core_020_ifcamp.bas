Attribute VB_Name = "func_core_020_ifcamp"
'-----------------------------------------------
'模块功能:
'   此模块用于放置zhazhupai006中的感应接口的控制
'   if=interface
'-----------------------------------------------
Public Sub if_IFonCo(WhichRange As String, WhichCountTimely As String, WhichCountBenchmark As String)
'程序功能：
'   eye接口，如果WhichCountTimely >= WhichCountBenchmark,对应的值为1，反之则为0
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
Dim page02 As String
Dim CountTimely As Integer
Dim CountBenchmark As Integer
Dim Range01 As String
page02 = cs_FV("ScreenSheet")
CountTimely = cs_FindValue(WhichCountTimely)
CountBenchmark = cs_FindValue(WhichCountBenchmark)
With Sheets(page02)
        .Range(WhichRange).Interior.Color = 6710784
        If CountTimely < CountBenchmark Then
        '如果有不同则指示灯为红色
            .Range(WhichRange).Interior.Color = 192
            .Range(WhichRange).Value = 0
        Else
        '相同为绿色
            .Range(WhichRange).Interior.Color = 5296274
            .Range(WhichRange).Value = 1
        End If
End With
End Sub

