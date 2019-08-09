VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Modifier des données de article"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4875
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click() 'OK

If OptionButton1.Value = True Then
    Call modifierArticles(1, TextBox1.Value)
ElseIf OptionButton2.Value = True Then
    Call modifierArticles(2, TextBox1.Value)
ElseIf OptionButton3.Value = True Then
    Call modifierArticles(3, TextBox1.Value)
ElseIf OptionButton4.Value = True Then
    Call modifierArticles(4, TextBox1.Value)
ElseIf OptionButton5.Value = True Then
    Call modifierArticles(5, TextBox1.Value)
ElseIf OptionButton6.Value = True Then
    Call modifierArticles(6, TextBox1.Value)
ElseIf OptionButton7.Value = True Then
    Call modifierArticles(7, TextBox1.Value)
ElseIf OptionButton8.Value = True Then
    Call modifierArticles(8, TextBox1.Value)
ElseIf OptionButton9.Value = True Then
    Call modifierArticles(9, TextBox1.Value)
ElseIf OptionButton10.Value = True Then
    Call modifierArticles(10, TextBox1.Value)
End If

End Sub

Private Sub CommandButton2_Click() 'Cancel

Unload Me

End Sub

Private Sub UserForm_Initialize()

Dim ligne As Integer
ligne = Selection.Row

'Initialiser
TextBox1.Value = Worksheets("PREPA SAP").Range("B" & ligne).Value

'Désélectionner les options
OptionButton1.Value = False
OptionButton2.Value = False
OptionButton3.Value = False

TextBox1.SetFocus

End Sub




