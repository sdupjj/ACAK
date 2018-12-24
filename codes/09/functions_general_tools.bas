Attribute VB_Name = "functions_general_tools"
'-----------------------------------------------
'模块功能:
'   此模块用于放置所有xlsm都可以用的程序
'-----------------------------------------------

Function num2asc2(ByVal n As Integer) As String
'程序功能：
'   输入数字返回英语字母
'程序版本：
'   1.0
'版本修订：
'   1.0 >>> 原始版本
    num2asc2 = Mid(Cells(1, n).Address, 2, IIf(n < 27, 1, 2))
End Function
