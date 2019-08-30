Attribute VB_Name = "connexionSAP"
Sub logonSAP()
'_________________________________________________________________________________________________'
                    'Logon SAP
'Variables
Dim SapGui, Applic, Connection, WSHShell
Dim identifiant As String, motDePasse As String, langue As String

'identifiant = "ng2b609"
'motDePasse = "Dr210591"
identifiant = "ng2b23d"
motDePasse = "RPS08201"

'identifiant = InputBox("Ecrivez votre identifiant de l'utilisateur", "RPS")
If StrPtr(identifiant) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenêtre
    MsgBox ("Vous avez annulé l'opération !")
    End 'Arrête tous les procedures en exécution
End If

'motDePasse = InputBox("Ecrivez votre mot de passe", "RPS")
If StrPtr(motDePasse) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenêtre
    MsgBox ("Vous avez annulé l'opération !")
    End
End If

langue = "FR"

Shell ("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

Set WSHShell = CreateObject("WScript.Shell")

Do Until WSHShell.AppActivate("SAP Logon") 'Attendre SAP ouvrir
    Application.Wait Now + TimeValue("0:00:01")
Loop

Set SapGui = GetObject("SAPGUI") 'get the interface of the SAPGUI object

If Not IsObject(SapGui) Then
    Exit Sub
End If

Set Applic = SapGui.GetScriptingEngine 'get the interface of the currently running SAP GUI process

If Not IsObject(Applic) Then
    Exit Sub
End If

Set Connection = Applic.openconnection("..SAP2000 Production             PGI")

If Not IsObject(Connection) Then
   Exit Sub
End If

Set session = Connection.Children(0)
If Connection.Children.Count < 1 Then
    Exit Sub
Else
    Set session = Connection.Children(0)
End If

If Not IsObject(session) Then
   Exit Sub
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = identifiant
session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = motDePasse

session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = langue
session.findById("wnd[0]").sendVKey 0

End Sub

Sub fermetureSAP()

session.findById("wnd[0]").Close
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

End Sub
