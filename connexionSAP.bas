Attribute VB_Name = "connexionSAP"
Option Explicit

Global session, wnd0, userArea, menuBar, statusBar, toolBar0

'Faire la connexion avec SAP, c'est-�-dire ouvrir et fermer une session

Sub logonSAP()
'_________________________________________________________________________________________________'
                    'Logon SAP
'Variables
Dim SapGui, Applic, Connection, WSHShell
Dim identifiant As String, motDePasse As String, langue As String

'Ouvrir logiciel
Shell ("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

Set WSHShell = CreateObject("WScript.Shell")

Do Until WSHShell.AppActivate("SAP Logon") 'Attendre SAP ouvrir
    Application.Wait Now + TimeValue("0:00:01")
Loop

'R�cup�rer l'interface de l'objet SAPGUI
Set SapGui = GetObject("SAPGUI")

If Not IsObject(SapGui) Then
    Exit Sub
End If

'R�cup�rer l'interface du processus SAP GUI en ex�cution
Set Applic = SapGui.GetScriptingEngine

If Not IsObject(Applic) Then
    Exit Sub
End If

'Connexion avec SAP PGI
Set Connection = Applic.openconnection("..SAP2000 Production             PGI")

If Not IsObject(Connection) Then
   Exit Sub
End If

'Session
Set session = Connection.Children(0)

If Connection.Children.Count < 1 Then
    Exit Sub
Else
    Set session = Connection.Children(0)
End If

If Not IsObject(session) Then
   Exit Sub
End If

'Demander les informations pour faire le login
connexion:
'identifiant = "ng2b609"
'motDePasse = "Dr210592"
identifiant = "ng2b23d"
motDePasse = "RPS08201"

'identifiant = InputBox("Ecrivez votre identifiant de l'utilisateur", "Connexion SAP")
If StrPtr(identifiant) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fen�tre
    MsgBox ("Vous avez annul� l'op�ration !")
    End 'Arr�te tous les procedures en ex�cution
End If

'motDePasse = InputBox("Ecrivez votre mot de passe", "Connexion SAP")
If StrPtr(motDePasse) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fen�tre
    MsgBox ("Vous avez annul� l'op�ration !")
    End
End If

langue = "FR"

'Variables
Dim messageErreur As String, userName As String

userName = session.info.user

Set wnd0 = session.findById("wnd[0]")
Set userArea = wnd0.findById("usr")
Set menuBar = wnd0.findById("mbar")
Set statusBar = wnd0.findById("sbar")
Set toolBar0 = wnd0.findById("tbar[0]")

'SAP R/3
wnd0.maximize
userArea.findById("txtRSYST-BNAME").Text = identifiant
userArea.findById("pwdRSYST-BCODE").Text = motDePasse
userArea.findById("txtRSYST-LANGU").Text = langue
wnd0.sendVKey 0 'Enter

'V�rification de la bonne connexion
If (statusBar.MessageType = "E") Then 'Erreur au connecter au SAP

    messageErreur = statusBar.Text
    Select Case MsgBox("La connexion SAP a �chou�e ! On a la message suivante : " & Chr(13) & "<<" & messageErreur _
    & ">>." & Chr(13) & "Voulez-vous ressayer la connexion ?", vbYesNo + vbExclamation, "Connexion �chou�e")
        Case vbYes
            GoTo connexion
        Case vbNo
            MsgBox ("Vous avez annul� l'op�ration !")
            wnd0.Close 'Fermer
            End
    End Select
    
End If

End Sub

Sub fermetureSAP()

session.findById("wnd[0]").Close 'Fermer
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press 'Confirmer la fermeture

End Sub
