Attribute VB_Name = "Module1"
Option Explicit 'rend obligatoire la declaration des variables avant leur utilisation

Sub supprimerArticles_SAP()
'_________________________________________________________________________________________________'
                    'Logon SAP
'Variables
Dim SapGui, Applic, Connection, session, WSHShell
Dim identifiant As String, motDePasse As String, langue As String

'On Error GoTo errHandler

'identifiant = "ng2b609"
'motDePasse = "Dr210591"
identifiant = "ng2b23d"
motDePasse = "RPS08201"

'identifiant = InputBox("Ecrivez votre identifiant de l'utilisateur", "RPS")
If StrPtr(identifiant) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
    Exit Sub
End If

'motDePasse = InputBox("Ecrivez votre mot de passe", "RPS")
If StrPtr(motDePasse) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
    Exit Sub
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

'_________________________________________________________________________________________________'
                    'Suppresion
Dim fichier As String, article As String, division As String, numeroMagasin As String, emplacement As String, dernier As String, i As Integer
Dim qteDemandee As String, supprimes As String, magasin As String, typeMagasin As String

fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
dernier = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row

'For i = 4 To dernier 'Les deux premieres lignes sont des exemples
For i = 19 To 19

    Workbooks(fichier).Activate
    article = ActiveSheet.Range("B" & i).Value 'CMS
    division = ActiveSheet.Range("J" & i).Value 'NTF ou (NZF)
    numeroMagasin = ActiveSheet.Range("L" & i).Value 'N18 ou (Z18)
    division = ActiveSheet.Range("J" & i).Value 'NTF ou (NZF)
    magasin = ActiveSheet.Range("K" & i).Value 'NENM ou (Z62M)
    typeMagasin = ActiveSheet.Range("M" & i).Value 'NEN ou (Z62)
    
    Debug.Print article
    Debug.Print division
    Debug.Print numeroMagasin

    '_________________________________________________________________________________________________'
                    'Modifier Article (Emplacement)
    emplacement = InputBox("Quel est le nouveau emplacement du article " & article & " ?")
    'emplacement = "0A0105"

    'session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
    session.findById("wnd[0]").sendVKey 0

    '-------- Modifier Article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
    
    'Configurer le niveau de organization (Nantes ou St Nazaire)
    session.findById("wnd[0]/tbar[1]/btn[6]").press 'ouvrir le "Niveaux de organization"
    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = "" 'Division
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = "" 'Magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").Text = "" 'Numero magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").Text = "" 'Type magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = division
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").Text = numeroMagasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").Text = typeMagasin

    session.findById("wnd[0]/tbar[0]/btn[0]").press

    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Modifier Article (MRP1, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    'session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Modifier Article (MRP2, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Gestion emplacements Masagin (CMS - CMS) --------
    session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text = emplacement
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    'session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

    '_________________________________________________________________________________________________'
                    'Transfert
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    'Recuper la quantité demandée
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "md04"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Etat dynamique des stocks actuel : écran initial --------
    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = article
    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").Text = division
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Etat dynamique des stocks --------
    qteDemandee = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG02[9,0]").Text
    Debug.Print qteDemandee
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    'Faire le transfert
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "lt01"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Créer ordre de transfert : écran initial --------
    session.findById("wnd[0]/usr/ctxtLTAK-LGNUM").Text = numeroMagasin 'Numéro de magasin
    session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").Text = "999" 'Code mouvement
    session.findById("wnd[0]/usr/ctxtLTAP-MATNR").Text = article 'Article
    session.findById("wnd[0]/usr/txtRL03T-ANFME").Text = qteDemandee 'Qté demandée
    session.findById("wnd[0]/usr/ctxtLTAP-WERKS").Text = division 'Division/Magasin
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    '-------- Créer ordre de transfert : générer poste OT --------
    
    [PROBLEMES !!!!!!!!!!!!!!!]
    
    session.findById("wnd[0]/tbar[0]/btn[0]").press 'Button 'Suite'
    session.findById("wnd[0]/tbar[0]/btn[0]").press 'Button 'Suite'
    session.findById("wnd[0]/tbar[0]/btn[0]").press 'Button 'Suite'
    
'    session.findById("wnd[0]/usr/txtRL03T-ANFME").Text = qteDemandee 'Qte demandee
'    session.findById("wnd[0]/usr/ctxtLTAP-VLTYP").Text = "A" 'De : Ty.
'    session.findById("wnd[0]/usr/ctxtLTAP-VLBER").Text = "B" 'De : A.S
'    session.findById("wnd[0]/usr/txtLTAP-VLPLA").Text = "C" 'De : Emplacem.
'    session.findById("wnd[0]/usr/ctxtLTAP-NLTYP").Text = "E" 'Prnt : Ty.
'    session.findById("wnd[0]/usr/ctxtLTAP-NLBER").Text = "F" 'Prnt : A.S
'    session.findById("wnd[0]/usr/txtLTAP-NLPLA").Text = "G" 'Prnt : Emplacem.
    
'    session.findById("wnd[0]").sendVKey 0
'    session.findById("wnd[0]").sendVKey 0
'    session.findById("wnd[0]").sendVKey 0
'    session.findById("wnd[0]/tbar[0]/btn[3]").press

    '_________________________________________________________________________________________________'
                    'Supprimer des articles

    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm06"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Position témoin suppresion article : écran de sélection --------
    session.findById("wnd[0]/usr/ctxtRM03G-MATNR").Text = article
    session.findById("wnd[0]/usr/ctxtRM03G-WERKS").Text = division
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Position témoin suppresion article : écran de données --------
    session.findById("wnd[0]/usr/chkRM03G-LVOMA").Selected = True 'Article
    session.findById("wnd[0]/usr/chkRM03G-LVOWK").Selected = True 'Division
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    'session.findById("wnd[0]/tbar[0]/btn[15]").press 'Terminer
    session.findById("wnd[0]/tbar[0]/btn[0]").press 'Suite
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    'Article supprimé
    supprimes = supprimes & article & " "

Next i

MsgBox ("La suppression des articles est fini." & Chr(13) & "Les articles suivants ont été supprimés : " & supprimes)

'Vider les cellules
'Workbooks(fichier).Activate
'ActiveSheet.Range("B4:I" & dernier).ClearContents
'ActiveSheet.Range("V4:V" & dernier).ClearContents

'Sauvegarder
Workbooks(fichier).Save

'Fermeture de la connexion
If MsgBox("La suppression des articles est fini. Voulez-vous fermer votre session SAP ?", vbYesNo, "RPS") = vbYes Then
    session.findById("wnd[0]").Close
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
End If
Exit Sub

errHandler:
    MsgBox "Une erreur est survenue !" & vbCrLf & "Numéro d'erreur : " & Err.Number & vbCrLf & "Description d'erreur : " & Err.Description
End Sub
