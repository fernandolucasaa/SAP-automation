Attribute VB_Name = "creerArticlesEnMasseSAP_PE1"
Option Explicit

Function fichierOuvert(fich As String) As Boolean

On Error Resume Next
If Workbooks(fich) Is Nothing Then
    fichierOuvert = False
Else
    fichierOuvert = True
End If
On Error GoTo 0 'Defaut

End Function

Sub creerArticlesEnMasse_SAPPE1()

Dim fichier As String
Dim premier As Integer, dernier As Integer, i As Integer, compteur As Integer
Dim ws As Worksheet

fichier = ThisWorkbook.Name

Windows(fichier).Activate
premier = 2
dernier = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
compteur = 0

Set ws = Windows(fichier).ActiveSheet

'Verification avant faire la création
Select Case MsgBox("Voulez-vous créer les articles suivants ?" & Chr(13) & Chr(13) & "Debut : " _
& ActiveSheet.Range("A" & premier) & Chr(13) & "Fin : " & ActiveSheet.Range("A" & dernier) _
, vbYesNo + vbQuestion, "Création des articles")
    Case vbNo
        MsgBox ("Vous avez annulé l'opération !")
        Exit Sub
End Select

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                'Créer des articles
            
'Faire une boucle pour créer l'article selectionné de chaque fichier
For i = premier To dernier

    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm01"
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (Ecran initial) --------
    session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").Key = "A" 'Branche : Construct. industrielles
    session.findById("wnd[0]/usr/cmbRMMG1-MTART").Key = "FATE" 'Type d'article : FORNITURE ATELIER
    session.findById("wnd[0]").sendVKey 0 'Enter

    'Sélection des vues (Division NZ01)
    session.findById("wnd[1]/tbar[0]/btn[19]").press 'Effacer la sélection
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Données de base 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'Achats
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).Selected = True 'Texte de commande d'achat
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).Selected = True 'Planification des besoins 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).Selected = True 'Planification des besoins 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).Selected = True 'Planification des besoins 3
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(14).Selected = True 'Planification des besoins 4
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 15 'Ajuster la position du scroll
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(16).Selected = True 'Données gén. divs./stockage 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(17).Selected = True 'Données gén. divs./stockage 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(21).Selected = True 'Comptabilité 1
    session.findById("wnd[1]/tbar[0]/btn[0]").press 'Enter
    
    'Nvx organisationnels
    Dim division As String, article As String
    
    division = ws.Range("D" & i).Value
    article = ws.Range("A" & i).Value
    
    session.findById("wnd[2]/usr/ctxtMARC-WERKS").Text = division
    session.findById("wnd[2]/usr/txtMARA-MFRPN").Text = article
    session.findById("wnd[2]/tbar[0]/btn[2]").press

    '-------- Créer article (Données de base 1) --------
    Dim designation As String, qteBase As String, grpeMarchand As String, statArt As String
  
    designation = ws.Range("F" & i).Value
    qteBase = ws.Range("G" & i).Value
    grpeMarchand = ws.Range("H" & i).Value
    statArt = ws.Range("I" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:1002/txtMAKT-MAKTX").Text = designation
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MEINS").Text = qteBase
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MATKL").Text = grpeMarchand
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MSTAE").Text = statArt
    session.findById("wnd[0]").sendVKey 0 'Enter

    '-------- Créer article (Achats) --------
    Dim grpAcheteurs As String, cleAchats As String, tempsReception As String
    
    grpAcheteurs = ws.Range("J" & i).Value
    cleAchats = ws.Range("K" & i).Value
    tempsReception = ws.Range("L" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/chkMARC-KAUTB").Selected = True 'Cde automatique
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = grpAcheteurs
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-MMSTA").Text = statArt
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2302/ctxtMARA-EKWSL").Text = cleAchats
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2303/txtMARC-WEBAZ").Text = tempsReception
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0 'Confirmer

    '-------- Créer article (Texte commande de achat) --------
    Dim texteCommande As String
    
    texteCommande = ws.Range("M" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = texteCommande
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (Planif. des besions 1) -------
    Dim codeABC As String, typePlan As String, ptCommande As String, grpPlanif As String, cleTailleLot As String, gestionnaire As String
    Dim tailleLotFixe As String
    
    grpPlanif = ws.Range("N" & i).Value
    codeABC = ws.Range("O" & i).Value
    typePlan = ws.Range("P" & i).Value
    ptCommande = ws.Range("Q" & i).Value
    gestionnaire = ws.Range("R" & i).Value
    cleTailleLot = ws.Range("S" & i).Value
    tailleLotFixe = ws.Range("T" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-DISGR").Text = grpPlanif
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-MAABC").Text = codeABC
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = typePlan
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = ptCommande
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text = gestionnaire
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = cleTailleLot
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/txtMARC-BSTFE").Text = tailleLotFixe
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (Planif. des besions 2) -------
    Dim typeApprov As String, magProduction As String, utilisationQuotas As String, magApproExt As String, delaiLivrais As String
    Dim cleHorizon As String
    
    typeApprov = ws.Range("U" & i).Value
    magProduction = ws.Range("V" & i).Value
    utilisationQuotas = ws.Range("W" & i).Value
    magApproExt = ws.Range("X" & i).Value
    delaiLivrais = ws.Range("Y" & i).Value
    cleHorizon = ws.Range("Z" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").Text = typeApprov
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").Text = magProduction
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-USEQU").Text = utilisationQuotas
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").Text = magApproExt
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").Text = delaiLivrais
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").Text = cleHorizon
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (Planif. des besions 3) -------
    Dim controleDispo As String
    
    controleDispo = ws.Range("AA" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = controleDispo
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Créer article (Planif. des besions 4) -------
    Dim indivCollectif As String
    
    indivCollectif = ws.Range("AB" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2495/ctxtMARC-SBDKZ").Text = indivCollectif
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (Donn.div./stockage 1) -------
    Dim emplacement As String
    
    emplacement = ws.Range("AC" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLZMGD1:2701/txtMARD-LGPBE").Text = emplacement
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (Donn.div./stockage 2) -------
    Dim centreProfit As String
    
    centreProfit = ws.Range("AD" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP20/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:5801/ctxtMARC-PRCTR").Text = centreProfit
    session.findById("wnd[0]").sendVKey 0

    '-------- Créer article (Comptabilité 1) -------
    Dim classeValoris As String, clValorCdeClt As String, clValProjet As String, prixStandard As String
    
    classeValoris = ws.Range("AE" & i).Value
    clValorCdeClt = ws.Range("AF" & i).Value
    clValProjet = ws.Range("AG" & i).Value
    prixStandard = ws.Range("AH" & i).Value
    
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-BKLAS").Text = classeValoris
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-EKLAS").Text = clValorCdeClt
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-QKLAS").Text = clValProjet
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/subSUBCURR:SAPLCKMMAT:0200/txtCKMMAT_DISPLAY-STPRS_1").Text = prixStandard
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0

    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]").sendVKey 0 'On retourne à l'ecran initial

    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour

    compteur = compteur + 1

Next i

'_________________________________________________________________________________________________'
                    'Fin de la création de l'article
                    
Windows(fichier).Activate
MsgBox ("La création des articles est finie ! Vous avez créé " & compteur & " articles !")

'Sauvegarder
Workbooks(fichier).Save

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

End Sub


