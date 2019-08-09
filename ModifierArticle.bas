Attribute VB_Name = "Module2"
Sub modifierButton()

UserForm1.Show

End Sub

Sub modifierArticles(optionChoisie As Integer, article As String)

Unload UserForm1 'Fermer

'_________________________________________________________________________________________________'
                    'Logon SAP
'Variables
Dim SapGui, Applic, Connection, session
Dim identifiant As String, motDePasse As String, langue As String

identifiant = "ng2b609"
motDePasse = "Dr210591"
'identifiant = "ng2b23d"
'motDePasse = "RPS08201"

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
'Vérification
If MsgBox("Avant continuer, vous avez déjà configurer le 'Niveaux Organization' pour Nantes ou Saint-Nazaire ?", vbYesNo + vbExclamation, "Niveaux de organization") = vbNo Then

    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
    session.findById("wnd[0]/tbar[1]/btn[6]").press 'ouvrir le "Niveaux de organization"
    MsgBox "Si le site n'est pas correc, changez les informations, sauvegardez la modification et quittez la session."
    Exit Sub
    
End If

'-------- Barre de recherche --------
session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
session.findById("wnd[0]").sendVKey 0

'-------- Modifier Article (Ecran initial) --------
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[0]").press 'Selection des vues

Dim valeur As String

If optionChoisie = 1 Then 'Designation
    
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    Dim designation As String
    designation = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text
    valeur = InputBox("La designation du article " & article & " est : " & designation & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    

ElseIf optionChoisie = 2 Then 'Texte de commande

     '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    Dim texteCommande As String
    texteCommande = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text
    valeur = InputBox("Le texte de commande du article " & article & " est : " & texteCommande & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

ElseIf optionChoisie = 3 Then 'Statut art. par div.

    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim statutArt As String
    statutArt = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text
    valeur = InputBox("Le statut art. par div. du article " & article & " est : " & statutArt & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

ElseIf optionChoisie = 4 Then 'Type planification
    
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim typePlanif As String
    typePlanif = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text
    valeur = InputBox("Le type de planification du article " & article & " est : " & typePlanif & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    If (valeur = "ND") Then
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
    End If
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

ElseIf optionChoisie = 5 Then 'Point de commande

    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim ptCommande As String
    ptCommande = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text
    valeur = InputBox("Le point de commande du article " & article & " est : " & ptCommande & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
         
ElseIf optionChoisie = 6 Then 'Valeur arrondie

    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim valeurArrondie As String
    valeurArrondie = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text
    valeur = InputBox("Le valeur arrondie du article " & article & " est : " & valeurArrondie & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

ElseIf optionChoisie = 7 Then 'Délai livrai

    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim delaiLivrai As String
    delaiLivrai = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text
    valeur = InputBox("Le delai livrais du article " & article & " est : " & delaiLivrai & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

ElseIf optionChoisie = 8 Then 'Clé calc. taille lot
    
     '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim cleCalcTailleLot As String
    cleCalcTailleLot = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text
    valeur = InputBox("La clé calc. taille lot du article " & article & " est : " & cleCalcTailleLot & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

ElseIf optionChoisie = 9 Then 'Numéro pce. fabricant

    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    Dim numFabricant As String
    numFabricant = session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text
    valeur = InputBox("Le numéro pce fabricant du article " & article & " est : " & numFabricant & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
ElseIf optionChoisie = 10 Then 'Emplacement
    
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP2, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    'session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Gestion emplacements Masagin (CMS - CMS) --------
    Dim emplacement As String
    emplacement = session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text
    valeur = InputBox("L'emplacement du article " & article & " est : " & emplacement & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text = valeur

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If

    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

End If

'Sauvegarder
'Workbooks(prepaPoint).Save

'Fermeture de la connexion
If MsgBox("La modification des articles est fini. Voulez-vous fermer votre session SAP ?", vbYesNo, "RPS") = vbYes Then
    session.findById("wnd[0]").Close
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
End If

End Sub


