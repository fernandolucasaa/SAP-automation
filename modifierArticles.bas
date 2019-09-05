Attribute VB_Name = "modifierArticles"
Option Explicit

'Modifier tous les articles de ce fichier, à chaque fois l'utilisateur doit choisir quelle modification à faire
'et quelle est la nouvelle valeur à mettre
'Modifier les articles pour Nantes et Saint-Nazaire

Sub modifierArticles2()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                    'Modifier Article
Dim article As String, fichier As String, i As Integer, fin As String, ref As String

fichier = ThisWorkbook.Name
fin = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row

Workbooks(fichier).Activate
ref = 4 'Reférence pour sauvegarder le niveau de organization
article = ActiveSheet.Range("B" & ref).Value

'-------- Barre de recherche --------
session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
session.findById("wnd[0]").sendVKey 0

'-------- Modifier Article (Ecran initial) --------
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article

'Modifier l'article pour le site à Nantes ou à Saint Nazaire
Dim division As String, magasin As String, numeroMagasin As String, typeMagasin As String, valeur As String
Dim typePlan As String, compteur As Integer

Workbooks(fichier).Activate
division = ActiveSheet.Range("J" & ref).Value 'NTF ou (NZF)
magasin = ActiveSheet.Range("K" & ref).Value 'NENM ou (Z62M)
numeroMagasin = ActiveSheet.Range("L" & ref).Value 'N18 ou (Z18)
typeMagasin = ActiveSheet.Range("M" & ref).Value 'NEN ou (Z62)

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
session.findById("wnd[1]/tbar[0]/btn[13]").press 'Sauvegarder comme paramétrage

session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour (F3)
session.findById("wnd[0]/tbar[0]/btn[3]").press

compteur = 0

For i = 4 To fin

    Workbooks(fichier).Activate
    article = ActiveSheet.Range("B" & i).Value
    
    Load UserForm1 'créer l'UserForm, mais pas l'afficher
    UserForm1.TextBox1 = article
    MsgBox ("Choisissez la modification à faire !")
    UserForm1.Show
    
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier Article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
    session.findById("wnd[0]/tbar[0]/btn[0]").press

    If UserForm1.OptionButton1 = True Then 'Designation
        
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        Dim designation As String
        designation = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text
        valeur = InputBox("La designation du article " & article & " est : " & designation & Chr(13) _
        & "Ecrivez la nouvelle designation : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text = valeur
        
        GoSub Enregistrer
        
    ElseIf UserForm1.OptionButton2 = True Then 'Texte de commande

        GoSub TexteDeCommande
        
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        Dim texteCommande As String
        texteCommande = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text
        valeur = InputBox("Le texte de commande du article " & article & " est : " & texteCommande & Chr(13) _
        & "Ecrivez le nouveau texte de commande : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = valeur
    
        GoSub Enregistrer

    ElseIf UserForm1.OptionButton3 = True Then 'Statut art. par div.

        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim statutArt As String
        statutArt = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text
        valeur = InputBox("Le statut art. par div. du article " & article & " est : " & statutArt & Chr(13) _
        & "Ecrivez le nouveau statut : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text = valeur
    
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If
        
    ElseIf UserForm1.OptionButton4 = True Then 'Type planification
    
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim typePlanif As String
        typePlanif = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text
        valeur = InputBox("Le type de planification du article " & article & " est : " & typePlanif & Chr(13) _
        & "Ecrivez le nouveau type : ")
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = valeur
    
        'Il faut verifier si on a la bonne clé pour le nouveau type de planification
        Dim cleCalcTailleLot As String
        cleCalcTailleLot = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text
        
        If (valeur = "VB" And cleCalcTailleLot = "") Then
            Select Case MsgBox("La clé calc. taille lot n'est pas la bonnne pour le nouveau type 'VB'. Il faut la modifier !" _
            & " Voulez-vous modifier pour 'EX' ?", vbYesNo, "RPS")
                Case vbYes
                    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = "EX"
            End Select
        End If
        
        If (valeur = "ND" And cleCalcTailleLot = "EX") Then
            Select Case MsgBox("La clé calc. taille lot n'est pas la bonnne pour le nouveau type 'ND'. Voulez-vous modifier pour '' ?", vbYesNo, "RPS")
                Case vbYes
                    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = ""
            End Select
        End If
        
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    ElseIf UserForm1.OptionButton5 = True Then 'Point de commande
    
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim ptCommande As String
        ptCommande = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text
        valeur = InputBox("Le point de commande du article " & article & " est : " & ptCommande & Chr(13) _
        & "Ecrivez le nouveau point : ")
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = valeur
        
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If
         
    ElseIf UserForm1.OptionButton6 = True Then 'Valeur arrondie
    
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim valeurArrondie As String
        valeurArrondie = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text
        valeur = InputBox("La valeur arrondie du article " & article & " est : " & valeurArrondie & Chr(13) _
        & "Ecrivez la nouvelle valeur : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text = valeur
        
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    ElseIf UserForm1.OptionButton7 = True Then 'Délai livrai
    
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim delaiLivrai As String
        delaiLivrai = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text
        valeur = InputBox("Le delai livrais du article " & article & " est : " & delaiLivrai & Chr(13) _
        & "Ecrivez le nouveau delai : ")
        session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text = valeur
        
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    ElseIf UserForm1.OptionButton8 = True Then 'Clé calc. taille lot
    
        GoSub MRP1
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim cleCalcTailleLot2 As String
        cleCalcTailleLot2 = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text
        valeur = InputBox("La clé calc. taille lot du article " & article & " est : " & cleCalcTailleLot2 & Chr(13) _
        & "Ecrivez la nouvelle clé : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = valeur
    
        typePlan = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text
    
        If (valeur = "" And typePlan = "VB") Then
            Select Case MsgBox("Le type planication n'est pas le bon pour la nouvelle clé ''. Il faut le modifier !" _
            & " Voulez-vous modifier pour 'ND' ?", vbYesNo, "RPS")
                Case vbYes
                    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND"
                    typePlan = "ND"
            End Select
        End If
        
        If (valeur = "EX" And typePlan = "ND") Then
            Select Case MsgBox("Le type planification n'est pas le bon pour la nouvelle clé 'EX'. Il faut le modifier !" _
            & " Voulez-vous modifier pour 'VB' ?", vbYesNo, "RPS")
                Case vbYes
                    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "VB"
                    typePlan = "VB"
            End Select
        End If
        
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    ElseIf UserForm1.OptionButton9 = True Then 'Numéro pce. fabricant
    
        GoSub Achats
        
        '-------- Modifier Article (Achats, CMS - CMS) --------
        Dim numFabricant As String
        numFabricant = session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text
        valeur = InputBox("Le numéro pce fabricant du article " & article & " est : " & numFabricant & Chr(13) _
        & "Ecrivez le nouveau numéro : ")
        session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text = valeur
    
        GoSub Enregistrer
    
    ElseIf UserForm1.OptionButton10 = True Then 'Emplacement
        
        GoSub DonneesGenDivStockage
        
        '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
        Dim emplacement As String
        emplacement = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").Text
        valeur = InputBox("L'emplacement du article " & article & " est : " & emplacement & Chr(13) _
        & "Ecrivez le nouveau emplacement : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").Text = valeur
        
        session.findById("wnd[0]/tbar[1]/btn[18]").press 'Continuer
        
        '-------- Modifier Article (Gestion emplacements Masagin, CMS - CMS) --------
        session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text = valeur
    
        GoSub Enregistrer

    ElseIf UserForm1.OptionButton11 = True Then 'Grp Acheteur
    
        GoSub Achats
        
        '-------- Modifier Article (Achats, CMS - CMS) --------
        Dim grpAcheteurs As String
        grpAcheteurs = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text
        valeur = InputBox("Le groupe acheteur du article " & article & " est : " & grpAcheteurs & Chr(13) _
        & "Ecrivez le nouveau groupe : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = valeur
        
        GoSub Enregistrer
    
    ElseIf UserForm1.OptionButton17 = True Then 'Gestionnaire
    
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim gestionnaire As String
        gestionnaire = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text
        valeur = InputBox("Le gestionnaire du article " & article & " est : " & gestionnaire & Chr(13) _
        & "Ecrivez le nouveau gestionnaire : ")
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text = valeur
        
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    ElseIf UserForm1.OptionButton12 = True Then 'Cle Horizon
    
        GoSub MRP1
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim cleHorizon As String
        cleHorizon = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text
        valeur = InputBox("La clé horizon du article " & article & " est : " & cleHorizon & Chr(13) _
        & "Ecrivez la nouvelle clé : ")
        session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text = valeur
        
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    ElseIf UserForm1.OptionButton13 = True Then 'Grp Marchandise
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        Dim grpMarchandise As String
        grpMarchandise = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MATKL").Text
        valeur = InputBox("Le groupe merchandise du article " & article & " est : " & grpMarchandise & Chr(13) _
        & "Ecrivez le nouveau groupe : ")
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MATKL").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            MsgBox ("Vous avez annulé l'opération ! La session SAP sera fermé !")
            Unload UserForm1
            fermetureSAP
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[1]/btn[18]").press 'Continuer
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARA-MATKL").Text = valeur
    
        GoSub Enregistrer
    
    ElseIf UserForm1.OptionButton14 = True Then 'Controle Dispo
    
        GoSub MRP2
        
        '-------- Modifier article (MRP 2, CMS - CMS) --------
        Dim controleDispo As String
        controleDispo = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text
        valeur = InputBox("Le controle disponibil. du article " & article & " est : " & controleDispo & Chr(13) _
        & "Ecrivez le nouveau controle : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = valeur
    
        GoSub Enregistrer

    ElseIf UserForm1.OptionButton15 = True Then 'Type magasin pour SM
    
        GoSub GestionEmplacementsMagasin
        
        '-------- Modifier Article (Gestion emplacements Masagin, CMS - CMS) --------
        Dim typeMagSM As String
        typeMagSM = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text
        valeur = InputBox("Le type magasin pour SM du article " & article & " est : " & typeMagSM & Chr(13) _
        & "Ecrivez le nouveau type : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text = valeur
        
        GoSub Enregistrer

    ElseIf UserForm1.OptionButton16 = True Then 'Type magasin EM
    
        GoSub GestionEmplacementsMagasin
        
        '-------- Modifier Article (Gestion emplacements Masagin, CMS - CMS) --------
        Dim typeMagEM As String
        typeMagEM = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text
        valeur = InputBox("Le type magasin EM du article " & article & " est : " & typeMagEM & Chr(13) _
        & "Ecrivez le nouveau type : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text = valeur
        
        GoSub Enregistrer
        
    End If
    
    Unload UserForm1
    compteur = compteur + 1
    
Next i

'Sauvegarder
Workbooks(fichier).Save

MsgBox ("Vous avez modifié " & compteur & " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

Exit Sub

Enregistrer:
    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        MsgBox ("Vous avez annulé l'opération ! La session SAP sera fermée !")
        Unload UserForm1
        fermetureSAP
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    Return

'[BUG]
Enregistrer2:
    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        MsgBox ("Vous avez annulé l'opération ! La session SAP sera fermée !")
        Unload UserForm1
        fermetureSAP
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    Return

Achats:
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return
    
TexteDeCommande:
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return

MRP1:
    GoSub TexteDeCommande
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return

MRP2:
    GoSub MRP1
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    'Si on a "ND", il y a une étape de plus, une message est affichée
    If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    End If
    
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return

DonneesGenDivStockage:
    GoSub MRP2
    
    '-------- Modifier Article (MRP2, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return

GestionEmplacementsMagasin:
    GoSub MRP2
    
    '-------- Modifier Article (MRP2, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return

End Sub
