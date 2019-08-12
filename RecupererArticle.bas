Attribute VB_Name = "Module1"
Option Explicit 'rend obligatoire la declaration des variables avant leur utilisation

Sub recupererArticles_SAP()
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
                    'Recuperer des donnees des articles
Dim fichier As String, article As String, dernier As String, i As Integer


fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
dernier = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row

'For i = 4 To dernier 'Les deux premieres lignes sont des exemples
For i = 4 To 4
    
    Workbooks(fichier).Activate
    article = ActiveSheet.Range("B" & i).Value 'CMS
    
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm03" 'Afficher des articles
    session.findById("wnd[0]").sendVKey 0

    '-------- Afficher article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article 'Article
    session.findById("wnd[0]/tbar[0]/btn[0]").press
     
    Dim division As String, magasin As String, numeroMagasin As String, typeMagasin As String

'
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
'
'
'    '-------- Afficher article (Données de base, CMS - CMS) --------
    Dim designation As String
    'designation = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text 'Désignation article
    
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
    
    ptCommande = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text 'Point de commande
    
    Workbooks(fichier).Activate
    ActiveSheet.Range("E" & i).Value = statutArt 'M1 ou vide
    ActiveSheet.Range("F" & i).Value = typePlanif 'ND ou VB
    ActiveSheet.Range("I" & i).Value = delaiLivrai
    ActiveSheet.Range("U" & i).Value = gestionnaire 'BF1 ou (CIG)
    ActiveSheet.Range("W" & i).Value = magasinProd 'NENM ou (Z62M)
    ActiveSheet.Range("X" & i).Value = magApproExt 'NENM ou (Z62M)
    'ActiveSheet.Range("V" & i).Value = cleCalcTailleLot 'EX ou (FX) ou vide
    ActiveSheet.Range("Z" & i).Value = cleHorizon 'N01 ou (Z01)

    ActiveSheet.Range("G" & i).Value = ptCommande

'    If (cleCalcTailleLot = "FX") Then 'St Nazaire
'        valeurArrondie = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTFE").Text '(Taille de lot fixe)
'    Else 'Nantes
'        valeurArrondie = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text 'Valeur arrondie
'    End If

    If (typePlanif = "VB") Then
        cleCalcTailleLot = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text 'Clé calc. taille lot
    End If
    Workbooks(fichier).Activate
    ActiveSheet.Range("H" & i).Value = valeurArrondie 'ou (Taille de lot fixe)

    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    '-------- Afficher article (MRP 2, CMS - CMS) --------
    Dim controleDispo As String, indivCollect As String
    
    controleDispo = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text 'Controle disponibil.
    indivCollect = session.findById("wnd[0]/usr/subSUB6:SAPLMGD1:2495/ctxtMARC-SBDKZ").Text 'Individuel/Collectif
    magasin = session.findById("wnd[0]/usr/subSUB1:SAPLMGD1:1005/ctxtRMMG1-LGORT").Text

    Workbooks(fichier).Activate
    ActiveSheet.Range("AB" & i).Value = controleDispo 'KP ou (O2)
    ActiveSheet.Range("AC" & i).Value = indivCollect '2
    ActiveSheet.Range("K" & i).Value = magasin 'NENM ou (Z62M)
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    '-------- Afficher article (Donnéees gén. div./stockage, CMS - CMS) --------
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
    ActiveSheet.Range("AG" & i).Value = classeValor '0510
    ActiveSheet.Range("AH" & i).Value = codePrix 'V
    ActiveSheet.Range("AI" & i).Value = basePrix '1
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'suivante

    'Retourner à l'accueil
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'buttom pour faire le retour
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'buttom pour faire le retour

Next i

End Sub



'Sub supprimerArticles_SAP()
''_________________________________________________________________________________________________'
'                    'Logon SAP
''Variables
'Dim SapGui, Applic, Connection, session, WSHShell
'Dim identifiant As String, motDePasse As String, langue As String
'
''On Error GoTo errHandler
'
''identifiant = "ng2b609"
''motDePasse = "Dr210591"
'identifiant = "ng2b23d"
'motDePasse = "RPS08201"
'
''identifiant = InputBox("Ecrivez votre identifiant de l'utilisateur", "RPS")
'If StrPtr(identifiant) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
'    Exit Sub
'End If
'
''motDePasse = InputBox("Ecrivez votre mot de passe", "RPS")
'If StrPtr(motDePasse) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
'    Exit Sub
'End If
'
'langue = "FR"
'
'Shell ("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
'
'Set WSHShell = CreateObject("WScript.Shell")
'
'Do Until WSHShell.AppActivate("SAP Logon") 'Attendre SAP ouvrir
'    Application.Wait Now + TimeValue("0:00:01")
'Loop
'
'Set SapGui = GetObject("SAPGUI") 'get the interface of the SAPGUI object
'
'If Not IsObject(SapGui) Then
'    Exit Sub
'End If
'
'Set Applic = SapGui.GetScriptingEngine 'get the interface of the currently running SAP GUI process
'
'If Not IsObject(Applic) Then
'    Exit Sub
'End If
'
'Set Connection = Applic.openconnection("..SAP2000 Production             PGI")
'
'If Not IsObject(Connection) Then
'   Exit Sub
'End If
'
'Set session = Connection.Children(0)
'If Connection.Children.Count < 1 Then
'    Exit Sub
'Else
'    Set session = Connection.Children(0)
'End If
'
'If Not IsObject(session) Then
'   Exit Sub
'End If
'
'session.findById("wnd[0]").maximize
'session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = identifiant
'session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = motDePasse
'
'session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = langue
'session.findById("wnd[0]").sendVKey 0
'
''_________________________________________________________________________________________________'
'
'Dim fichier As String, article As String, division As String, numeroMagasin As String, emplacement As String, dernier As String, i As Integer
'Dim qteDemandee As String, supprimes As String
'
'fichier = ThisWorkbook.Name
'
'Workbooks(fichier).Activate
'dernier = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
'
''For i = 4 To dernier 'Les deux premieres lignes sont des exemples
'For i = 19 To 19
'
'    Workbooks(fichier).Activate
'    article = ActiveSheet.Range("B" & i).Value 'CMS
'    division = ActiveSheet.Range("J" & i).Value 'NTF ou (NZF)
'    numeroMagasin = ActiveSheet.Range("L" & i).Value 'N18 ou (Z18)
'
'    Debug.Print article
'    Debug.Print division
'    Debug.Print numeroMagasin
'
'    '_________________________________________________________________________________________________'
'                    'Modifier Article (Emplacement)
'    emplacement = InputBox("Quel est le nouveau emplacement du article " & article & " ?")
'    'emplacement = "0A0105"
'
'    'session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
'
'    '-------- Barre de recherche --------
'    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
'    session.findById("wnd[0]").sendVKey 0
'
'    '-------- Modifier Article (Ecran initial) --------
'    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
'
'    session.findById("wnd[0]/tbar[0]/btn[0]").press
'    'session.findById("wnd[0]/tbar[0]/btn[0]").press 'selection de vues
'
'    '-------- Modifier Article (Données de base, CMS - CMS) --------
'    session.findById("wnd[0]/tbar[1]/btn[18]").press
'
'    '-------- Modifier Article (Achats, CMS - CMS) --------
'    session.findById("wnd[0]/tbar[1]/btn[18]").press
'
'    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
'    session.findById("wnd[0]/tbar[1]/btn[18]").press
'
'    '-------- Modifier Article (MRP1, CMS - CMS) --------
'    session.findById("wnd[0]/tbar[1]/btn[18]").press
'    'session.findById("wnd[0]/tbar[1]/btn[18]").press
'
'    '-------- Modifier Article (MRP2, CMS - CMS) --------
'    session.findById("wnd[0]/tbar[1]/btn[18]").press
'
'    '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
'    session.findById("wnd[0]/tbar[1]/btn[18]").press
'
'    '-------- Gestion emplacements Masagin (CMS - CMS) --------
'    session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text = emplacement
'    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
'    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
'    'session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
'
'    '_________________________________________________________________________________________________'
'                    'Transfert
'    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
'
'    'Recuper la quantité demandée
'    '-------- Barre de recherche --------
'    session.findById("wnd[0]/tbar[0]/okcd").Text = "md04"
'    session.findById("wnd[0]").sendVKey 0
'
'    '-------- Etat dynamique des stocks actuel : écran initial --------
'    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = article
'    session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").Text = division
'    session.findById("wnd[0]").sendVKey 0
'
'    '-------- Etat dynamique des stocks --------
'    qteDemandee = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG02[9,0]").Text
'    Debug.Print qteDemandee
'    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
'    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
'
'    'Faire le transfert
'    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
'
'    '-------- Barre de recherche --------
'    session.findById("wnd[0]/tbar[0]/okcd").Text = "lt01"
'    session.findById("wnd[0]").sendVKey 0
'
'    '-------- Créer ordre de transfert : écran initial --------
'    session.findById("wnd[0]/usr/ctxtLTAK-LGNUM").Text = numeroMagasin 'Numéro de magasin
'    session.findById("wnd[0]/usr/ctxtLTAK-BWLVS").Text = "999" 'Code mouvement
'    session.findById("wnd[0]/usr/ctxtLTAP-MATNR").Text = article 'Article
'    session.findById("wnd[0]/usr/txtRL03T-ANFME").Text = qteDemandee 'Qté demandée
'    session.findById("wnd[0]/usr/ctxtLTAP-WERKS").Text = division 'Division/Magasin
'    session.findById("wnd[0]/tbar[0]/btn[0]").press
'
'    '-------- Créer ordre de transfert : générer poste OT --------
'
'    [PROBLEMES !!!!!!!!!!!!!!!]
'
'    session.findById("wnd[0]/tbar[0]/btn[0]").press 'Button 'Suite'
'    session.findById("wnd[0]/tbar[0]/btn[0]").press 'Button 'Suite'
'    session.findById("wnd[0]/tbar[0]/btn[0]").press 'Button 'Suite'
'
''    session.findById("wnd[0]/usr/txtRL03T-ANFME").Text = "1" 'Qte demandee
''    session.findById("wnd[0]/usr/ctxtLTAP-VLTYP").Text = "A" 'De : Ty.
''    session.findById("wnd[0]/usr/ctxtLTAP-VLBER").Text = "B" 'De : A.S
''    session.findById("wnd[0]/usr/txtLTAP-VLPLA").Text = "C" 'De : Emplacem.
''    session.findById("wnd[0]/usr/ctxtLTAP-NLTYP").Text = "E" 'Prnt : Ty.
''    session.findById("wnd[0]/usr/ctxtLTAP-NLBER").Text = "F" 'Prnt : A.S
''    session.findById("wnd[0]/usr/txtLTAP-NLPLA").Text = "G" 'Prnt : Emplacem.
'
''    session.findById("wnd[0]").sendVKey 0
''    session.findById("wnd[0]").sendVKey 0
''    session.findById("wnd[0]").sendVKey 0
''    session.findById("wnd[0]/tbar[0]/btn[3]").press
'
'    '_________________________________________________________________________________________________'
'                    'Supprimer des articles
'
'    '-------- Barre de recherche --------
'    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm06"
'    session.findById("wnd[0]").sendVKey 0
'
'    '-------- Position témoin suppresion article : écran de sélection --------
'    session.findById("wnd[0]/usr/ctxtRM03G-MATNR").Text = article
'    session.findById("wnd[0]/usr/ctxtRM03G-WERKS").Text = division
'    session.findById("wnd[0]").sendVKey 0
'
'    '-------- Position témoin suppresion article : écran de données --------
'    session.findById("wnd[0]/usr/chkRM03G-LVOMA").Selected = True 'Article
'    session.findById("wnd[0]/usr/chkRM03G-LVOWK").Selected = True 'Division
'
'    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
'    'session.findById("wnd[0]/tbar[0]/btn[15]").press 'Terminer
'    session.findById("wnd[0]/tbar[0]/btn[0]").press 'Suite
'    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
'
'    'Article supprimé
'    supprimes = supprimes & article & " "
'
'Next i
'
'MsgBox ("La suppression des articles est fini." & Chr(13) & "Les articles suivants ont été supprimés : " & supprimes)
'
''Vider les cellules
''Workbooks(fichier).Activate
''ActiveSheet.Range("B4:I" & dernier).ClearContents
''ActiveSheet.Range("V4:V" & dernier).ClearContents
'
''Sauvegarder
''Workbooks(fichier).Save
'
''Fermeture de la connexion
'If MsgBox("La suppression des articles est fini. Voulez-vous fermer votre session SAP ?", vbYesNo, "RPS") = vbYes Then
'    session.findById("wnd[0]").Close
'    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'End If
'Exit Sub
'
'errHandler:
'    MsgBox "Une erreur est survenue !" & vbCrLf & "Numéro d'erreur : " & Err.Number & vbCrLf & "Description d'erreur : " & Err.Description
'End Sub
