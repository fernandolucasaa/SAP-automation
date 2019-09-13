VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "Modifier des données de article"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click() 'OK

Me.Hide

End Sub

Private Sub CommandButton2_Click() 'Cancel

Unload Me
MsgBox ("Vous avez annulé l'opération ! La session SAP sera fermé !")
fermetureSAP
End

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
OptionButton14.Value = False
OptionButton15.Value = False
OptionButton16.Value = False
OptionButton17.Value = False

TextBox1.SetFocus

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If (CloseMode = vbformcontrlmenu) Then 'Finir l'opération si fermer le formulaire
    MsgBox ("Vous avez annulé l'opération ! La session SAP sera fermée !")
    fermetureSAP
    End
End If

End Sub
