VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Modifier des données de article"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   OleObjectBlob   =   "UserForm1_28-08.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click() 'OK

Me.Hide

End Sub

Private Sub CommandButton2_Click() 'Cancel

Unload Me

End Sub

Private Sub UserForm_Initialize()

'Désélectionner les options
OptionButton1.Value = False
OptionButton2.Value = False
OptionButton3.Value = False
OptionButton4.Value = False
OptionButton5.Value = False
OptionButton6.Value = False
OptionButton7.Value = False
OptionButton8.Value = False
OptionButton9.Value = False
OptionButton10.Value = False
OptionButton11.Value = False
OptionButton12.Value = False
OptionButton13.Value = False

TextBox1.SetFocus

End Sub




