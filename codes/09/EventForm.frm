VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EventForm 
   Caption         =   "Event"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7455
   OleObjectBlob   =   "EventForm.frx":0000
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "EventForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ListBox1_Click()
    itemnum = ListBox1.ListIndex
    Label1.Caption = Sheets("core_actions").Range("C" & (itemnum + 2)).Value
End Sub
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
        itemnum = ListBox1.ListIndex
        Cells(Val(Label4.Caption), Val(Label5.Caption)) = ListBox1.List(itemnum, 1)
'        Sheets("core_screen").Cells(Val(Label4.Caption), Val(Label5.Caption)) = ListBox1.List(itemnum, 1)
        Unload EventForm
End Sub

