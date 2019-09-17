Attribute VB_Name = "creerArticles"
Option Explicit

Global session

'Cr�er tous les article du fichier. L'utilisateur doit confirmer la bonne cr�ation des articles pour
'les premiers n articles cr��s
'Cr�er des articles pour Nantes et Saint-Nazaire

Sub creerArticles_SAP()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                    'Creer une article
Dim fichier As String, article As String, ws As Worksheet
Dim fin As Integer, compteur As Integer, limite As Integer, i As Integer

fichier = ThisWorkbook.Name
Set ws = Windows(fichier).ActiveSheet

fin = ws.Cells(Rows.Count, 2).End(xlUp).Row
compteur = 0 'qt� totale de articles cr�es
limite = 5 'limite de v�rification

'On Error GoTo 0
'Pour d�bugger le code il faut mettre la ligne en bas en commentaire
On Error GoTo errHandler

For i = 4 To fin

    '-------- Barre de recherche --------
    toolBar0.findById("okcd").Text = "mm01"
    wnd0.sendVKey 0 'Enter
    
    '-------- Cr�er article (Ecran initial) --------
    Dim modele As String
    
    Workbooks(fichier).Activate
    modele = ActiveSheet.Range("A" & i).Value '8MODELNENM ou (8MODELZ62M)
    article = ActiveSheet.Range("B" & i).Value
    
    'V�rification du CMS
    If (Len(article) <> 10) Then
    
        MsgBox "La taille de l'article " & article & " est incorrecte !" & Chr(13) & "L'article se trouve " _
        & "dans la ligne " & i & " , fixez la valeur et relancez l'op�ration !", vbExclamation, "Erreur CMS"
        MsgBox ("La session SAP sera ferm�� !")
        fermetureSAP
        Exit Sub
        
    End If
    
    userArea.findById("ctxtRMMG1-MATNR").Text = article  'Article
    userArea.findById("cmbRMMG1-MBRSH").Key = "M"  'Branche
    userArea.findById("cmbRMMG1-MTART").Key = "CMS"  'Type d'article (CMS - CMS)
    userArea.findById("ctxtRMMG1_REF-MATNR").Text = modele  'Mod�le

    'Cr�er l'article pour le site � Nantes ou � Saint Nazaire
    Dim division As String, magasin As String, numeroMagasin As String, typeMagasin As String
    
    Workbooks(fichier).Activate
    division = ActiveSheet.Range("J" & i).Value 'NTF ou (NZF)
    magasin = ActiveSheet.Range("K" & i).Value 'NENM ou (Z62M)
    numeroMagasin = ActiveSheet.Range("L" & i).Value 'N18 ou (Z18)
    typeMagasin = ActiveSheet.Range("M" & i).Value 'NEN ou (Z62)

    session.findById("wnd[0]/tbar[1]/btn[6]").press 'ouvrir le "Niveaux de organization"
    
    'Configurer le niveau de organization
    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = "" 'Division
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = "" 'Magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").Text = "" 'Numero magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").Text = "" 'Type magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = division
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").Text = numeroMagasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").Text = typeMagasin
    session.findById("wnd[1]/tbar[0]/btn[5]").press 'S�lection des vues

    'S�lection des vues
    session.findById("wnd[1]/tbar[0]/btn[19]").press 'Demarquer tout
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Donn�es de base
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'Achats
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Texte de commande
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'MRP 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'MRP 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).Selected = True 'Donn�es g�n. div./stockage
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).Selected = True 'Gestion emplacements magasin
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(15).Selected = True 'Comptabilit�
    session.findById("wnd[1]/tbar[0]/btn[0]").press 'Suite
    
    'V�rification si le CMS utilis� est d�j� cr��
    verifierCMS

    '-------- Cr�er article (Donn�es de base, CMS - CMS) --------
    Dim designation As String
    
    Workbooks(fichier).Activate
    designation = ActiveSheet.Range("C" & i).Value
    
    'V�rification de la designation
    If verifierEntree(designation, "designation", article, i) = False Then
        MsgBox ("La session SAP sera ferm�� !")
        fermetureSAP
        Exit Sub
    End If
    
    userArea.findById("subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text = designation 'D�signation article
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant

    '-------- Cr�er article (Achats, CMS - CMS) --------
    Dim grpAcheteurs As String, tempsReception As String, numFabricant As String
    
    Workbooks(fichier).Activate
    grpAcheteurs = ActiveSheet.Range("R" & i).Value 'BF1 ou (CIG)
    tempsReception = ActiveSheet.Range("Y" & i).Value '2
    numFabricant = ActiveSheet.Range("AJ" & i).Value
    
    userArea.findById("subSUB2:SAPLMGD1:2301/chkMARC-KAUTB").Selected = True 'Cde automatique
    userArea.findById("subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = grpAcheteurs 'Groupe d'acheteurs
    userArea.findById("subSUB4:SAPLMGD1:2303/txtMARC-WEBAZ").Text = tempsReception 'Temps de r�ception
    userArea.findById("subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text = numFabricant 'N� pce fabricant
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant
    wnd0.sendVKey 0 'Enter

    '-------- Cr�er article (Texte de commande, CMS - CMS) --------
    Dim texteCommande As String
    
    Workbooks(fichier).Activate
    texteCommande = ActiveSheet.Range("D" & i).Value
    
    'V�rification du texte de commande
    If verifierEntree(texteCommande, "texte de commande", article, i) = False Then
        MsgBox ("La session SAP sera ferm�� !")
        fermetureSAP
        Exit Sub
    End If

    userArea.findById("subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = texteCommande 'Texte de commande
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant

    '-------- Cr�er article (MRP 1, CMS - CMS) --------
    Dim statutArt As String, typePlanif As String, ptCommande As String, valeurArrondie As String, delaiLivrai As String
    Dim gestionnaire As String, magasinProd As String, magApproExt As String, cleCalcTailleLot As String, cleHorizon As String
    
    Workbooks(fichier).Activate
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

    userArea.findById("subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text = statutArt 'Statut art. par div.
    userArea.findById("subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = typePlanif 'Type planification
    userArea.findById("subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = ptCommande 'Point de commande
    userArea.findById("subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text = gestionnaire 'Gestionnaire

    'Nantes : valuer arrondie, St Nazaire : taille de lot fixe
    If (cleCalcTailleLot = "FX") Then 'St Nazaire
        userArea.findById("subSUB4:SAPLMGD1:2483/txtMARC-BSTFE").Text = valeurArrondie '(Taille de lot fixe)
    Else 'Nantes
        userArea.findById("subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text = valeurArrondie 'Valeur arrondie
    End If

    'Nantes : type 'VB', cle 'EX', St Nazaire : type 'VB', cle 'FX'
    If (typePlanif = "VB") Then
        userArea.findById("subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = cleCalcTailleLot 'Cl� calc. taille lot
    End If

    userArea.findById("subSUB6:SAPLMGD1:2484/ctxtMARC-LGPRO").Text = magasinProd 'Magasin production
    userArea.findById("subSUB6:SAPLMGD1:2484/ctxtMARC-LGFSB").Text = magApproExt 'Mag. pour appro. ext
    userArea.findById("subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text = delaiLivrai 'D�lai pr�v. livrais
    userArea.findById("subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text = cleHorizon 'Cl� d'horizon
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant
    wnd0.sendVKey 0 'Enter

    '-------- Cr�er article (MRP 2, CMS - CMS) --------
    Dim controleDispo As String, indivCollect As String

    Workbooks(fichier).Activate
    controleDispo = ActiveSheet.Range("AB" & i).Value 'KP ou (02)
    indivCollect = ActiveSheet.Range("AC" & i).Value '2

    If (division = "NTF") Then 'Nantes, le control disponibil. pour St Nazaire est deja rempli
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = controleDispo 'Controle disponibil.
    End If

    userArea.findById("subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").caretPosition = 2 'ligne ajout�e, car il avait de bug quand VB
    wnd0.sendVKey 0 'ligne ajout�e, car il avait de bug quand VB
    userArea.findById("subSUB6:SAPLMGD1:2495/ctxtMARC-SBDKZ").Text = indivCollect 'Individuel/Collectif
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant

    '-------- Cr�er article (Donn�ees g�n. div./stockage, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant

    '-------- Cr�eer article (Gestion emplacements magasin, CMS - CMS) --------
    Dim typeMagSM As String, typeMagEM As String

    Workbooks(fichier).Activate
    typeMagSM = ActiveSheet.Range("AE" & i).Value 'NEN ou (Z62)
    typeMagEM = ActiveSheet.Range("AF" & i).Value 'NEN ou (Z62)

    userArea.findById("subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text = typeMagSM 'Type magasin pour SM
    userArea.findById("subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text = typeMagEM 'Type magasin EM
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant

    '-------- Cr�er article (Comptabilit�, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[26]").press
    userArea.findById("subSUB3:SAPLMGD1:2802/ctxtMBEW-BKLAS").Text = "0510" 'Classe valorisation
    userArea.findById("subSUB3:SAPLMGD1:2802/ctxtMBEW-BKLAS").caretPosition = 4
    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Ecran suivant
    wnd0.sendVKey 0
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press 'Quittez le traitement et sauvegarder

    'Articles cr�ees
    compteur = compteur + 1

    'Retourner � l'accueil
    toolBar0.findById("btn[3]").press 'buttom pour faire le retour
    toolBar0.findById("btn[3]").press 'buttom pour faire le retour
    
    'V�rification manuelle de l'utilisateur
    If (compteur = limite) Then
        
        MsgBox "Vous avez cr�� " & limite & " articles. V�rifiez si les articles sont corrects dans le SAP." _
        & " Apr�s finir votre v�rification, laissez votre session SAP ouverte dans l'�cran initial !", vbExclamation, _
        "Verifiez des articles"
        Select Case MsgBox("Voulez-vous continuer la cr�ation des articles ?", vbYesNo + vbQuestion, "Continuer op�ration")
            Case vbNo
                Exit For
        End Select
        
    End If

Next i

'Cr�ation termin�e
MsgBox ("La cr�ation des articles est finie ! Vous avez cr�e " & compteur & " articles.")

'Sauvegarder
Workbooks(fichier).Save

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

Exit Sub

errHandler:

    MsgBox "Un erreur a �t� trouv�!" & Chr(13) & Chr(13) _
    & "Num�ro de l'erreur :        " & Err.Number & Chr(13) & Chr(13) _
    & "Description de l'erreur :   " & Err.Description & Chr(13) & Chr(13) _
    & "Status bar du SAP :         " & statusBar.Text, vbExclamation, "Erreur"
    MsgBox "La procedure est finie !", vbExclamation
    Exit Sub

    Resume

End Sub

'V�rifier si num�ro CMS est d�j� cr��
Sub verifierCMS()

If session.ActiveWindow.Type = "GuiModalWindow" Then 'fen�tre pop-up du SAP
    
    MsgBox "Une fen�tre pop-up s'affiche avec la message suivante : " & Chr(13) _
    & "<< " & session.ActiveWindow.PopupDialogText & " >>" & Chr(13) _
    & "La session SAP sera ferm�e !", vbExclamation, "Erreur"
    session.findById(session.ActiveWindow.Name).Close 'wnd[2]
    session.findById(session.ActiveWindow.Name).Close 'wnd[1]
    fermetureSAP
    End
    
End If

End Sub

'V�rifier les messages d'erreur dans le status bar
Sub verifierErreur()

Dim messageErreur As String

If (statusBar.MessageType = "E") Then

    messageErreur = statusBar.Text
    MsgBox ("L'erreur suivant a �t� cr�� : " & Chr(13) & "<<" & messageErreur & ">>." & Chr(13) _
    & "La session SAP sera ferm�e !")
    fermetureSAP
    End
    
End If

End Sub

'V�rifier la bonne designation et texte de commande
Function verifierEntree(valeur As String, variable As String, article As String, pos As Integer) As Boolean

verifierEntree = True

If (valeur <> UCase(valeur)) Then
        
    MsgBox variable & "de l'article " & article & " doit �tre en majuscule !" & Chr(13) & "L'article se trouve " _
    & "dans la ligne " & pos & " , fixez la valeur et relancez l'op�ration !", vbExclamation, "Erreur"
    verifierEntree = False
    
End If
    
If (Len(valeur) > 40) Then
    
    MsgBox "Le nombre des caract�res de " & variable & " de l'article " & article & " est superior � 40 !" & Chr(13) & "L'article se trouve " _
    & "dans la ligne " & pos & " , fixez la valeur et relancez l'op�ration !", vbExclamation, "Erreur"
    verifierEntree = False
    
End If

End Function
