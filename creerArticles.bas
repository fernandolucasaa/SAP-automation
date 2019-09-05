Attribute VB_Name = "creerArticles"
Option Explicit

'Créer tous les article du fichier. L'utilisateur doit confirmer la bonne création des articles à
'chaque n creations
'Créer des articles pour Nantes et Saint-Nazaire

Sub creerArticles_SAP()

'Se connecter au SAP
logonSAP

'_________________________________________________________________________________________________'
                    'Creer une article
Dim fichier As String, article As String, modele As String, designation As String, i As Integer
Dim fin As Integer, nouveaux As String, compteur As Integer, cpt As Integer, limite As Integer

fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
fin = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
compteur = 0 'qté totale de articles crées
cpt = 0 'qté avant de demander une vérification
limite = 5 'demander à l'utilisateur de vérifier la création à chaque fois qu'on crée cette qté

For i = 4 To fin 'Les deux premieres lignes sont des exemples

    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm01"
    session.findById("wnd[0]").sendVKey 0

    Workbooks(fichier).Activate
    modele = ActiveSheet.Range("A" & i).Value '8MODELNENM ou (8MODELZ62M)
    article = ActiveSheet.Range("B" & i).Value
    designation = ActiveSheet.Range("C" & i).Value

    '-------- Créer article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article 'Article
    session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").Key = "M" 'Branche
    session.findById("wnd[0]/usr/cmbRMMG1-MTART").Key = "CMS" 'Type d'article (CMS - CMS)
    session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").Text = modele 'Modèle

    'Créer l'article pour le site à Nantes ou à Saint Nazaire
    Dim division As String, magasin As String, numeroMagasin As String, typeMagasin As String

    Workbooks(fichier).Activate
    division = ActiveSheet.Range("J" & i).Value 'NTF ou (NZF)
    magasin = ActiveSheet.Range("K" & i).Value 'NENM ou (Z62M)
    numeroMagasin = ActiveSheet.Range("L" & i).Value 'N18 ou (Z18)
    typeMagasin = ActiveSheet.Range("M" & i).Value 'NEN ou (Z62)

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
    session.findById("wnd[1]/tbar[0]/btn[5]").press

    'Effacer la selection
    session.findById("wnd[1]/tbar[0]/btn[19]").press

    'Sélection des vues
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Données de base
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'Achats
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Texte de commande
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'MRP 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'MRP 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).Selected = True 'Données gén. div./stockage
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).Selected = True 'Gestion emplacements magasin
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(15).Selected = True 'Comptabilité
    'session.findById("wnd[1]/tbar[0]/btn[0]").press 'Retour à la fenetre "Niveaux de organization"
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    '-------- Créer article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text = designation 'Désignation article
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Créer article (Achats, CMS - CMS) --------
    Workbooks(fichier).Activate
    Dim grpAcheteurs As String, tempsReception As String, numFabricant As String
    grpAcheteurs = ActiveSheet.Range("R" & i).Value 'BF1 ou (CIG)
    tempsReception = ActiveSheet.Range("Y" & i).Value '2
    numFabricant = ActiveSheet.Range("AJ" & i).Value

    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/chkMARC-KAUTB").Selected = True 'Cde automatique
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = grpAcheteurs 'Groupe d'acheteurs
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2303/txtMARC-WEBAZ").Text = tempsReception 'Temps de réception
    session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text = numFabricant 'N° pce fabricant
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (Texte de commande, CMS - CMS) --------
    Workbooks(fichier).Activate
    Dim texteCommande As String
    texteCommande = ActiveSheet.Range("D" & i).Value

    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = texteCommande 'Texte de commande
    'session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").setSelectionIndexes 6, 6
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Créer article (MRP 1, CMS - CMS) --------
    Workbooks(fichier).Activate
    Dim statutArt As String, typePlanif As String, ptCommande As String, valeurArrondie As String, delaiLivrai As String
    Dim gestionnaire As String, magasinProd As String, magApproExt As String, cleCalcTailleLot As String, cleHorizon As String

    statutArt = ActiveSheet.Range("E" & i).Value 'M1 ou vide
    typePlanif = ActiveSheet.Range("F" & i).Value 'ND ou VB
    ptCommande = ActiveSheet.Range("G" & i).Value
    valeurArrondie = ActiveSheet.Range("H" & i).Value 'ou (Taille de lot fixe)
    delaiLivrai = ActiveSheet.Range("I" & i).Value
    gestionnaire = ActiveSheet.Range("U" & i).Value 'BF1 ou (CIG)
    magasinProd = ActiveSheet.Range("W" & i).Value 'NENM ou (Z62M)
    magApproExt = ActiveSheet.Range("X" & i).Value 'NENM ou (Z62M)
    cleCalcTailleLot = ActiveSheet.Range("V" & i).Value 'EX ou (FX) ou vide
    cleHorizon = ActiveSheet.Range("Z" & i).Value 'N01 ou (Z01)

    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text = statutArt 'Statut art. par div.
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = typePlanif 'Type planification
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = ptCommande 'Point de commande
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text = gestionnaire 'Gestionnaire

    If (cleCalcTailleLot = "FX") Then 'St Nazaire
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTFE").Text = valeurArrondie '(Taille de lot fixe)
    Else 'Nantes
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text = valeurArrondie 'Valeur arrondie
    End If

    If (typePlanif = "VB") Then
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = cleCalcTailleLot 'Clé calc. taille lot
    End If

    session.findById("wnd[0]/usr/subSUB6:SAPLMGD1:2484/ctxtMARC-LGPRO").Text = magasinProd 'Magasin production
    session.findById("wnd[0]/usr/subSUB6:SAPLMGD1:2484/ctxtMARC-LGFSB").Text = magApproExt 'Mag. pour appro. ext
    session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text = delaiLivrai 'Délai prév. livrais
    session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text = cleHorizon 'Clé d'horizon
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (MRP 2, CMS - CMS) --------
    Workbooks(fichier).Activate
    Dim controleDispo As String, indivCollect As String

    controleDispo = ActiveSheet.Range("AB" & i).Value 'KP ou (02)
    indivCollect = ActiveSheet.Range("AC" & i).Value '2

    If (division = "NTF") Then 'Nantes, le control disponibil. pour St Nazaire est deja rempli
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = controleDispo 'Controle disponibil.
    End If

    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").caretPosition = 2 'ligne ajoutée, car il avait de bug quand VB
    session.findById("wnd[0]").sendVKey 0 'ligne ajoutée, car il avait de bug quand VB
    session.findById("wnd[0]/usr/subSUB6:SAPLMGD1:2495/ctxtMARC-SBDKZ").Text = indivCollect 'Individuel/Collectif
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Créer article (Donnéees gén. div./stockage, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Créeer article (Gestion emplacements magasin, CMS - CMS) --------
    Workbooks(fichier).Activate
    Dim typeMagSM As String, typeMagEM As String

    typeMagSM = ActiveSheet.Range("AE" & i).Value 'NEN ou (Z62)
    typeMagEM = ActiveSheet.Range("AF" & i).Value 'NEN ou (Z62)

    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text = typeMagSM 'Type magasin pour SM
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text = typeMagEM 'Type magasin EM
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Créer article (Comptabilité, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[26]").press
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2802/ctxtMBEW-BKLAS").Text = "0510" 'Classe valorisation
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2802/ctxtMBEW-BKLAS").caretPosition = 4
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

    'Article créee
    nouveaux = nouveaux & article & " "
    compteur = compteur + 1
    cpt = cpt + 1

    'Retourner à l'accueil
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'buttom pour faire le retour
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'buttom pour faire le retour
    
    'Vérification manuelle de l'utilisateur
    If (cpt = limite) Then
        
        'cpt = 0
        MsgBox "Vous avez créé " & limite & " articles. Vérifiez si les articles sont corrects dans le SAP." _
        & " Aprés finir votre vérification, laissez votre session SAP ouverte dans l'écran initial !", vbExclamation, _
        "Verifiez des articles"
        Select Case MsgBox("Voulez-vous continuer la création des articles ?", vbYesNo + vbQuestion, "Continuer opération")
            Case vbNo
                Exit For
        End Select
        
    End If

Next i

MsgBox ("La création des articles est finie ! Vous avez crée " & compteur & " articles.")

'Vider les cellules
'Workbooks(fichier).Activate
'ActiveSheet.Range("B4:I" & fin).ClearContents
'ActiveSheet.Range("V4:V" & fin).ClearContents

'Sauvegarder
Workbooks(fichier).Save

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

End Sub