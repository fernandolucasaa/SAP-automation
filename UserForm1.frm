VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Modifier des données de article"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   OleObjectBlob   =   "UserForm1.frx":0000
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
MsgBox ("Vous avez annulé l'opération ! La session SAP sera fermé !")
fermetureSAP
End

End Sub

Private Sub UserForm_Initialize()

Dim ctrl As Control

'Désélectionner les options
For Each ctrl In UserForm1.Controls
    If TypeName(ctrl) = "CheckBox" Then
        ctrl.Value = False
    End If
Next ctrl

TextBox1.SetFocus

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If (CloseMode = vbformcontrlmenu) Then 'Finir l'opération si fermer le formulaire
    MsgBox ("Vous avez annulé l'opération ! La session SAP sera fermée !")
    fermetureSAP
    End
End If

End Sub
