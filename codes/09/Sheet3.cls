VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Selection.Row() >= 2 And (Selection.Column() = 4 Or Selection.Column() = 8) Then
        EventForm.Label4.Caption = Selection.Row()
        EventForm.Label5.Caption = Selection.Column()
        EventForm.Show
    End If
End Sub
