Attribute VB_Name = "connexionSAP_PE1"
Option Explicit

Global session

'Faire la connexion avec SAP, c'est-à-dire ouvrir et fermer une session

Sub logonSAP()
'_________________________________________________________________________________________________'
                    'Logon SAP
'Variables
Dim SapGui, Applic, Connection, WSHShell

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

Set Connection = Applic.openconnection("...000-NEW ERP - ECC6 PE1 Production")

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

End Sub

Sub fermetureSAP()

session.findById("wnd[0]").Close 'Fermer
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press 'Confirmer la fermeture

End Sub

