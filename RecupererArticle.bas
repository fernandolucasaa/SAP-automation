Attribute VB_Name = "Module1"
Option Explicit 'rend obligatoire la declaration des variables avant leur utilisation

Sub recupererArticles_SAP()
'_________________________________________________________________________________________________'
                    'Logon SAP
'Variables
Dim SapGui, Applic, Connection, session, WSHShell
Dim identifiant As String, motDePasse As String, langue As String

'On Error GoTo errHandler

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
                    'Recuperer des donnees des articles
Dim fichier As String, article As String, dernier As String, i As Integer

fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
dernier = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row

For i = 4 To dernier 'Les deux premieres lignes sont des exemples

    Workbooks(fichier).Activate
    article = ActiveSheet.Range("B" & i).Value 'CMS
    
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm03" 'Afficher des articles
    session.findById("wnd[0]").sendVKey 0

    '-------- Afficher article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article 'Article
    session.findById("wnd[0]/tbar[0]/btn[0]").press
     
    Dim division As String, magasin As String, numeroMagasin As String, typeMagasin As String

    'session.findById("wnd[0]/tbar[0]/btn[0]").press

'    'Configurer le niveau de organization (Nantes ou St Nazaire)
'    session.findById("wnd[0]/tbar[1]/btn[6]").press 'ouvrir le "Niveaux de organization"
'    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = "" 'Division
'    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = "" 'Magasin
'    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").Text = "" 'Numero magasin
'    session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").Text = "" 'Type magasin
'    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = division
'    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = magasin
'    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").Text = numeroMagasin
'    session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").Text = typeMagasin
'    'session.findById("wnd[1]").sendVKey 4
'    'session.findById("wnd[0]/tbar[1]/btn[5]").press 'ouvrir la selection des vues
'
'    'Effacer la selection
'    session.findById("wnd[1]/tbar[0]/btn[19]").press
'
'    'Sélection des vues
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Données de base
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'Achats
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Texte de commande
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'MRP 1
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'MRP 2
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).Selected = True 'Données gén. div./stockage
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).Selected = True 'Gestion emplacements magasin
'    session.findById("wnd[1]/tbar[0]/btn[0]").press


    '-------- Afficher article (Données de base, CMS - CMS) --------
    Dim designation As String
    designation = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text 'Désignation article
    
    Dim uniteBase As String, grpMarchandise As String
    
    uniteBase = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MEINS").Text
    grpMarchandise = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MATKL").Text
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("C" & i).Value = designation
    ActiveSheet.Range("P" & i).Value = uniteBase 'PCE
    ActiveSheet.Range("Q" & i).Value = grpMarchandise 'Q224
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    '-------- Afficher article (Achats, CMS - CMS) --------
    Dim grpAcheteurs As String, tempsReception As String, numFabricant As String

    grpAcheteurs = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text  'Groupe d'acheteurs
    tempsReception = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2303/txtMARC-WEBAZ").Text 'Temps de réception
    numFabricant = session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text  'N° pce fabricant
    division = session.findById("wnd[0]/usr/subSUB1:SAPLMGD1:1001/ctxtRMMG1-WERKS").Text
    
    Dim cleValeurAchat As String
    cleValeurAchat = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2302/ctxtMARA-EKWSL").Text
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("R" & i).Value = grpAcheteurs 'BF1 ou (CIG)
    ActiveSheet.Range("Y" & i).Value = tempsReception '2
    ActiveSheet.Range("AJ" & i).Value = numFabricant
    ActiveSheet.Range("J" & i).Value = division 'NTF ou (NZF)
    ActiveSheet.Range("T" & i).Value = cleValeurAchat 'Z001
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    '-------- Afficher article (Texte de commande, CMS - CMS) --------
    Dim texteCommande As String
    texteCommande = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("D" & i).Value = texteCommande
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante
    
    '-------- Afficher article (MRP 1, CMS - CMS) --------
    Dim statutArt As String, typePlanif As String, ptCommande As String, valeurArrondie As String, delaiLivrai As String
    Dim gestionnaire As String, magasinProd As String, magApproExt As String, cleCalcTailleLot As String, cleHorizon As String
    
    statutArt = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text 'Statut art. par div.
    typePlanif = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text 'Type planification
    gestionnaire = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text 'Gestionnaire
    magasinProd = session.findById("wnd[0]/usr/subSUB6:SAPLMGD1:2484/ctxtMARC-LGPRO").Text 'Magasin production
    magApproExt = session.findById("wnd[0]/usr/subSUB6:SAPLMGD1:2484/ctxtMARC-LGFSB").Text 'Mag. pour appro. ext
    delaiLivrai = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text 'Délai prév. livrais
    cleHorizon = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text 'Clé d'horizon
    cleCalcTailleLot = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text 'Clé calc. taille lot
    
    ptCommande = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text 'Point de commande
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("E" & i).Value = statutArt 'M1 ou vide
    ActiveSheet.Range("F" & i).Value = typePlanif 'ND ou VB
    ActiveSheet.Range("I" & i).Value = delaiLivrai
    ActiveSheet.Range("U" & i).Value = gestionnaire 'BF1 ou (CIG)
    ActiveSheet.Range("W" & i).Value = magasinProd 'NENM ou (Z62M)
    ActiveSheet.Range("X" & i).Value = magApproExt 'NENM ou (Z62M)
    ActiveSheet.Range("V" & i).Value = cleCalcTailleLot 'EX ou (FX) ou vide
    ActiveSheet.Range("Z" & i).Value = cleHorizon 'N01 ou (Z01)

    ActiveSheet.Range("G" & i).Value = ptCommande

    If (cleCalcTailleLot = "FX") Then 'St Nazaire
        valeurArrondie = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTFE").Text '(Taille de lot fixe)
    Else 'Nantes
        valeurArrondie = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text 'Valeur arrondie
    End If

    Workbooks(fichier).Activate
    ActiveSheet.Range("H" & i).Value = valeurArrondie 'ou (Taille de lot fixe)

    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    '-------- Afficher article (MRP 2, CMS - CMS) --------
    Dim controleDispo As String, indivCollect As String, indicateurPeriode As String
    
    controleDispo = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text 'Controle disponibil.
    indivCollect = session.findById("wnd[0]/usr/subSUB6:SAPLMGD1:2495/ctxtMARC-SBDKZ").Text 'Individuel/Collectif
    magasin = session.findById("wnd[0]/usr/subSUB1:SAPLMGD1:1005/ctxtRMMG1-LGORT").Text
    indicateurPeriode = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2491/ctxtMARC-PERKZ").Text
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("AB" & i).Value = "'" & controleDispo 'KP ou (02)
    ActiveSheet.Range("AC" & i).Value = indivCollect '2
    ActiveSheet.Range("K" & i).Value = magasin 'NENM ou (Z62M)
    ActiveSheet.Range("AA" & i).Value = indicateurPeriode 'M
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    '-------- Afficher article (Donnéees gén. div./stockage, CMS - CMS) --------
    Dim indicPeriode As String
    
    indicPeriode = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2702/ctxtMARA-IPRKZ").Text
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("AD" & i).Value = indicPeriode 'J
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    '-------- Afficher article (Gestion emplacements magasin, CMS - CMS) --------
    Dim typeMagSM As String, typeMagEM As String
    
    typeMagSM = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text 'Type magasin pour SM
    typeMagEM = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text 'Type magasin EM
    numeroMagasin = session.findById("wnd[0]/usr/subSUB1:SAPLMGD1:1006/ctxtRMMG1-LGNUM").Text
    typeMagasin = session.findById("wnd[0]/usr/subSUB1:SAPLMGD1:1006/ctxtRMMG1-LGTYP").Text
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("AE" & i).Value = typeMagSM 'NEN ou (Z62)
    ActiveSheet.Range("AF" & i).Value = typeMagEM 'NEN ou (Z62)
    ActiveSheet.Range("L" & i).Value = numeroMagasin 'N18 ou (Z18)
    ActiveSheet.Range("M" & i).Value = typeMagasin 'NEN ou (Z62)
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    '-------- Afficher article (Comptabilité, CMS - CMS) --------
    Dim classeValor As String, codePrix As String, basePrix As String
    classeValor = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2802/ctxtMBEW-BKLAS").Text
    codePrix = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2802/ctxtMBEW-VPRSV").Text
    basePrix = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2802/txtMBEW-PEINH").Text
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("AG" & i).Value = "'" & classeValor '0510
    ActiveSheet.Range("AH" & i).Value = codePrix 'V
    ActiveSheet.Range("AI" & i).Value = basePrix '1
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    'Retourner à l'accueil
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'buttom pour faire le retour
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'buttom pour faire le retour

Next i

'Sauvegarder
'Workbooks(fichier).Save

'Fermeture de la connexion
If MsgBox("La récupération des articles est fini. Voulez-vous fermer votre session SAP ?", vbYesNo, "RPS") = vbYes Then
    session.findById("wnd[0]").Close
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
End If
Exit Sub

errHandler:
    MsgBox "Une erreur est survenue !" & vbCrLf & "Numéro d'erreur : " & Err.Number & vbCrLf & "Description d'erreur : " & Err.Description
End Sub
