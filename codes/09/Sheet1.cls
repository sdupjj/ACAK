VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    PageName02 = cs_FV("ScreenSheet")
    ColS = cs_FV("M02 Col Start Number")
    ColE = ColS + cs_FV("M02 Cols") - 1
    RowSS = cs_FV("M02 Row Start Number")
    RowE = RowSS + cs_FV("M02 Rows") - 1
    If (Selection.Row() >= RowSS And Selection.Row() <= RowE) And (Selection.Column() >= ColS And Selection.Column() <= ColE) Then
        EventForm.Label4.Caption = Selection.Row()
        EventForm.Label5.Caption = Selection.Column()
        EventForm.Show
    End If
End Sub

