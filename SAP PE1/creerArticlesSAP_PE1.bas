Attribute VB_Name = "creerArticlesSAP_PE1"
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

Sub creerArticles_SAPPE1()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                'Créer des articles
            
Dim fichierDevis As String, fichier As String, chemin As String
Dim premier As Integer, dernier As Integer, i As Integer, compteur As Integer, j As Integer

chemin = "\\Mycorp.corp\dfs\ProdControlLogistics\Services\RPS_FEB\Devis"
fichier = ThisWorkbook.Name

Windows(fichier).Activate
premier = Selection.Row
dernier = premier + Selection.Rows.Count - 1
compteur = 0

'Faire une boucle pour créer tous les articles selectionnés de chaque fichier
For i = premier To dernier

    Windows(fichier).Activate
    fichierDevis = ActiveSheet.Range("A" & i).Value + ".xlsm" 'nom du fichier

    'Ouvrir fichier DEVIS
    If (fichierOuvert(fichierDevis) = False) Then
        Workbooks.Open Filename:=chemin & "\" & fichierDevis
    End If

    'Calculer combien des feuils "FATE" il y a en chaque fichier
    Dim qteFeuils As Integer, ws As Worksheet
    qteFeuils = 0
    
    For Each ws In Workbooks(fichierDevis).Worksheets
        If InStr(ws.Name, "FATE_") <> 0 Then
            qteFeuils = qteFeuils + 1
        End If
    Next ws
    
    'Faire une boucle pour créer tous les articles d'un même fichier
    For j = 1 To qteFeuils
    
        '-------- Barre de recherche --------
        session.findById("wnd[0]/tbar[0]/okcd").Text = "mm01"
        session.findById("wnd[0]").sendVKey 0
    
        '-------- Créer article (Ecran initial) --------
        session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").Key = "A" 'Branche : Construct. industrielles
        session.findById("wnd[0]/usr/cmbRMMG1-MTART").Key = "FATE" 'Type d'article : FORNITURE ATELIER
        session.findById("wnd[0]").sendVKey 0 'Enter
    
        'Sélection des vues
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
        
        division = "NZ01"
        article = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("D28").Value
        
        session.findById("wnd[2]/usr/ctxtMARC-WERKS").Text = division
        session.findById("wnd[2]/usr/txtMARA-MFRPN").Text = article
        session.findById("wnd[2]/tbar[0]/btn[2]").press
    
        '-------- Créer article (Données de base 1) --------
        Dim designation As String, qteBase As String, grpeMarchand As String, statArt As String, articleReparable As String
      
        designation = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("C20").Value 'Designation SAP
        qteBase = "PCE"
        grpeMarchand = "Q224"
        articleReparable = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("I19").Value 'Oui ou non
        
        If (articleReparable = "OUI") Then
            statArt = "ZR"
        Else 'non
            statArt = "Z5"
        End If
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:1002/txtMAKT-MAKTX").Text = designation
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MEINS").Text = qteBase
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MATKL").Text = grpeMarchand
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MSTAE").Text = statArt
        session.findById("wnd[0]").sendVKey 0 'Enter
    
        '-------- Créer article (Achats) --------
        Dim grpAcheteurs As String, cleAchats As String, tempsReception As String
        
        grpAcheteurs = "T0A"
        cleAchats = "C010"
        tempsReception = "3"
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/chkMARC-KAUTB").Selected = True 'Cde automatique
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = grpAcheteurs
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-MMSTA").Text = statArt
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2302/ctxtMARA-EKWSL").Text = cleAchats
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2303/txtMARC-WEBAZ").Text = tempsReception
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0 'Confirmer
    
        '-------- Créer article (Texte commande de achat) --------
        Dim texteCommande As String, fornisseur As String, equipBase As String, equipCompl As String, reference As String
        
        designation = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("C20").Value 'Designation SAP
        fornisseur = "FOURNISSEUR " + Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("I25").Value
        equipBase = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("C24").Value
        equipCompl = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("C25").Value
        reference = "FEB " + Mid(fichierDevis, 7, 11)
        texteCommande = designation + vbCr + "" + vbCr + fornisseur + vbCr + "" + vbCr + "SAP " + equipBase + " / " + equipCompl + vbCr + "" + vbCr _
        + reference + vbCr + ""
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = texteCommande
        session.findById("wnd[0]").sendVKey 0
    
        '-------- Créer article (Planif. des besions 1) -------
        Dim codeABC As String, typePlan As String, ptCommande As String, grpPlanif As String, cleTailleLot As String, gestionnaire As String
        Dim tailleLotFixe As String
        
        grpPlanif = "G01S"
        codeABC = "C"
        
        If statArt = "ZR" Then
            typePlan = "ND"
        Else 'Z5
            typePlan = "VB"
        End If
        
        ptCommande = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("I20").Value
        gestionnaire = "T0A"
        cleTailleLot = "ZX"
        tailleLotFixe = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("I21").Value
        
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
        
        typeApprov = "F"
        magProduction = "5RM"
        utilisationQuotas = "Z"
        magApproExt = "5RM"
        delaiLivrais = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("I27").Value
        cleHorizon = "F05"
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").Text = typeApprov
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").Text = magProduction
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-USEQU").Text = utilisationQuotas
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").Text = magApproExt
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").Text = delaiLivrais
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").Text = cleHorizon
        session.findById("wnd[0]").sendVKey 0
    
        '-------- Créer article (Planif. des besions 3) -------
        Dim controleDispo As String
        
        controleDispo = "02"
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = controleDispo
        session.findById("wnd[0]").sendVKey 0
        
        '-------- Créer article (Planif. des besions 4) -------
        Dim indivCollectif As String
        
        indivCollectif = "2"
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2495/ctxtMARC-SBDKZ").Text = indivCollectif
        session.findById("wnd[0]").sendVKey 0
    
        '-------- Créer article (Donn.div./stockage 1) -------
        Dim emplacement As String
        
        emplacement = "CREA_FATE"
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLZMGD1:2701/txtMARD-LGPBE").Text = emplacement
        session.findById("wnd[0]").sendVKey 0
    
        '-------- Créer article (Donn.div./stockage 2) -------
        Dim centreProfit As String
        
        centreProfit = "FR10COMM"
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP20/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:5801/ctxtMARC-PRCTR").Text = centreProfit
        session.findById("wnd[0]").sendVKey 0
    
        '-------- Créer article (Comptabilité 1) -------
        Dim classeValoris As String, clValorCdeClt As String, clValProjet As String, prixStandard As String
        
        classeValoris = "02"
        clValorCdeClt = "02"
        clValProjet = "02"
        prixStandard = Workbooks(fichierDevis).Worksheets("FATE_" & j).Range("I26").Value
        
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-BKLAS").Text = classeValoris
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-EKLAS").Text = clValorCdeClt
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/ctxtMBEW-QKLAS").Text = clValProjet
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB1:SAPLCKMMAT:0010/tabsTABS/tabpPPLF/ssubSUBML:SAPLCKMMAT:0100/subSUBCURR:SAPLCKMMAT:0200/txtCKMMAT_DISPLAY-STPRS_1").Text = prixStandard
        session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]").sendVKey 0

        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[11]").press
    '    session.findById("wnd[0]/tbar[0]/btn[11]").press
    '    session.findById("wnd[0]").sendVKey 0
    '    session.findById("wnd[0]").sendVKey 0
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
        compteur = compteur + 1
    
    Next j
    
    Workbooks(fichierDevis).Close SaveChanges:=False
    
Next i

'_________________________________________________________________________________________________'
                    'Fin de la création de l'article

MsgBox ("La création des articles est finie ! Vous avez créé " + compteur + " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

End Sub
