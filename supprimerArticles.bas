Attribute VB_Name = "supprimerArticles"
Option Explicit

'Faire la suppression des articles en trois étapes : Emplacement, Transfert, Suppression

Sub supprimerArticles_SAP()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                    'Suppression des articles
Dim fichier As String, dernier As String, i As Integer, compteur As String

fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
dernier = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
compteur = 0 'qté des articles suprimés

'For i = 4 To dernier
For i = 8 To 8

    'Variables
    Dim article As String, division As String, numeroMagasin As String, emplacement As String
    Dim qteDemandee As String, magasin As String, typeMagasin As String

    Workbooks(fichier).Activate
    article = ActiveSheet.Range("B" & i).Value 'CMS
    division = ActiveSheet.Range("J" & i).Value 'NTF ou (NZF)
    numeroMagasin = ActiveSheet.Range("L" & i).Value 'N18 ou (Z18)
    division = ActiveSheet.Range("J" & i).Value 'NTF ou (NZF)
    magasin = ActiveSheet.Range("K" & i).Value 'NENM ou (Z62M)
    typeMagasin = ActiveSheet.Range("M" & i).Value 'NEN ou (Z62)
    'article = "8405033596"
    
    '_________________________________________________________________________________________________'
                    'Modifier Article (Emplacement)
    'emplacement = InputBox("Quel est le nouveau emplacement du article " & article & " ?")
    emplacement = "0A0105"

    '-------- Barre de recherche --------
    toolBar0.findById("okcd").Text = "mm02"
    wnd0.sendVKey 0 'Enter

    '-------- Modifier Article (Ecran initial) --------
    userArea.findById("ctxtRMMG1-MATNR").Text = article
    
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
    session.findById("wnd[1]/tbar[0]/btn[0]").press 'Suite
    
    'Aller jusqu'à fenêtre MRP1
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant (Données de base)
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant (Achats)
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant (Texte de commande)

    '-------- Modifier Article (MRP1, CMS - CMS) --------
    If (userArea.findById("subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
        session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant
    End If
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant (MRP2)
    
    '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
    userArea.findById("subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").Text = emplacement
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant (Données gén. div./stockage)

    '-------- Gestion emplacements Masagin (CMS - CMS) --------
    userArea.findById("subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text = emplacement
    toolBar0.findById("btn[11]").press 'Sauvegarder (Retouner à l'ecran initial)
    toolBar0.findById("btn[3]").press 'Retour

    '_________________________________________________________________________________________________'
                    'Transfert
                    
    'Récupérer la quantité demandée :
    
    '-------- Barre de recherche --------
    toolBar0.findById("okcd").Text = "md04"
    wnd0.sendVKey 0 'Enter
    
    '-------- Etat dynamique des stocks actuel : écran initial --------
    userArea.findById("tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = article
    userArea.findById("tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").Text = division
    wnd0.sendVKey 0 'Enter
    
    '-------- Etat dynamique des stocks --------
    qteDemandee = userArea.findById("subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG02[9,0]").Text
    Debug.Print qteDemandee
    toolBar0.findById("btn[3]").press 'Retour
    toolBar0.findById("btn[3]").press 'Retour
    
    'Faire le transfert :
    
    '-------- Barre de recherche --------
    toolBar0.findById("okcd").Text = "lt01"
    wnd0.sendVKey 0 'Enter
    
    '-------- Créer ordre de transfert : écran initial --------
    userArea.findById("ctxtLTAK-LGNUM").Text = numeroMagasin 'Numéro de magasin
    userArea.findById("ctxtLTAK-BWLVS").Text = "999" 'Code mouvement
    userArea.findById("ctxtLTAP-MATNR").Text = article 'Article
    userArea.findById("txtRL03T-ANFME").Text = qteDemandee 'Qté demandée
    userArea.findById("ctxtLTAP-WERKS").Text = division 'Division/Magasin
    toolBar0.findById("btn[0]").press 'Touche "Suite"
    
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
    toolBar0.findById("okcd").Text = "mm06"
    wnd0.sendVKey 0 'Enter
    
    '-------- Position témoin suppresion article : écran de sélection --------
    userArea.findById("ctxtRM03G-MATNR").Text = article
    userArea.findById("ctxtRM03G-WERKS").Text = division
    wnd0.sendVKey 0 'Enter
    
    '-------- Position témoin suppresion article : écran de données --------
    userArea.findById("chkRM03G-LVOMA").Selected = True 'Article
    userArea.findById("chkRM03G-LVOWK").Selected = True 'Division
    
    toolBar0.findById("btn[11]").press 'Sauvegarder
    toolBar0.findById("btn[0]").press 'Suite
    toolBar0.findById("btn[3]").press 'Retour
    
    'Articles supprimés
    compteur = compteur + 1

Next i

'Suppresion terminéé
MsgBox ("La suppression des articles est finie !" & Chr(13) & "Vous avez supprimé " & compteur & " articles.")

'Sauvegarder
Workbooks(fichier).Save

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

'Exit Sub
'
'errHandler:
'    MsgBox "Une erreur est survenue !" & vbCrLf & "Numéro d'erreur : " & Err.Number & vbCrLf & "Description d'erreur : " & Err.Description
End Sub






