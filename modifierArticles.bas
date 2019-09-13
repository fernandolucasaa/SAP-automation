Attribute VB_Name = "modifierArticles"
Option Explicit

'Modifier tous les articles de ce fichier, � chaque fois l'utilisateur doit choisir quelle modification � faire
'et quelle est la nouvelle valeur � mettre
'Modifier les articles pour Nantes et Saint-Nazaire

Sub modifierArticles2()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                    'Modifier Article
Dim article As String, fichier As String, i As Integer, fin As String, ref As String

fichier = ThisWorkbook.Name
fin = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
fin = 6

Workbooks(fichier).Activate
ref = 4 'Ref�rence pour sauvegarder le niveau de organization
article = ActiveSheet.Range("B" & ref).Value

'-------- Barre de recherche --------
session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
session.findById("wnd[0]").sendVKey 0

'-------- Modifier Article (Ecran initial) --------
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article

'Modifier l'article pour le site � Nantes ou � Saint Nazaire
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
session.findById("wnd[1]/tbar[0]/btn[5]").press 'S�lection des vues

'Selection des vues
session.findById("wnd[1]/tbar[0]/btn[19]").press 'Demarquer tout
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Donn�es de base
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).Selected = True 'Achats
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).Selected = True 'Texte de commande
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).Selected = True 'MRP 1
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'MRP 2
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Donn�es g�n. div./stockage
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'Gestion emplacements magasin
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'Comptabilit�
session.findById("wnd[1]/tbar[0]/btn[14]").press 'Sauvegarder comme param�trage
session.findById("wnd[1]").Close 'Fermer fen�tre

session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour (F3)

compteur = 0

Load UserForm1 'cr�er l'UserForm, mais pas l'afficher
UserForm1.TextBox1 = article
Workbooks(fichier).Activate
UserForm1.TextBox2 = ActiveSheet.Range("B" & fin).Value
MsgBox ("Choisissez les modifications � faire pour tous les articles!")
UserForm1.Show

'Si l'utilisateur n'est pas choisi des options
Dim ctrl As Control, flag As Integer
flag = False

For Each ctrl In UserForm1.Controls
    If TypeName(ctrl) = "CheckBox" Then
        If ctrl.Value = True Then
            flag = True
        End If
    End If
Next ctrl

If flag = False Then
    MsgBox ("Vous n'avez pas choisi faire des modifications ! La session SAP sera ferm�e !")
    Unload UserForm1
    fermetureSAP
    Exit Sub
End If

'Demander des entrees
Const qteOptions As Integer = 17
Dim entrees(1 To 17) As String, variable(1 To 17) As String

variable(1) = "de la designation"
variable(2) = "du texte de commande"
variable(3) = "du statut art. par div."
variable(4) = "du type de planification"
variable(5) = "de la designation"
variable(6) = "du texte de commande"
variable(7) = "du delai livrais"
variable(8) = "de la cl� calc. taille lot"
variable(9) = "de la designation"
variable(10) = "de l'emplacement"
variable(11) = "du groupe acheteur"
variable(12) = "du gestionnaire"
variable(13) = "de la cl� horizon"
variable(14) = "du groupe merchandise"
variable(15) = "du controle Dispo"
variable(16) = "du type magasin pour SM"
variable(17) = "du type magasin EM"

For i = 1 To 17
    
    If UserForm1.Controls("CheckBox" & i).Value = True Then
        entrees(i) = InputBox("Ecrivez la nouvelle valeur " + variable(i) + " pour les articles :")
    End If

Next i


'Boucle pour faire des modifications
For i = 4 To fin

    Workbooks(fichier).Activate
    article = ActiveSheet.Range("B" & i).Value
    
'    Load UserForm1 'cr�er l'UserForm, mais pas l'afficher
'    UserForm1.TextBox1 = article
'    MsgBox ("Choisissez la modification � faire !")
'    UserForm1.Show
    
'    '-------- Barre de recherche --------
'    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
'    session.findById("wnd[0]").sendVKey 0
'
'    '-------- Modifier Article (Ecran initial) --------
'    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
'    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    'V�rifier quelles modification � faire
    If UserForm1.CheckBox1 = True Then 'Designation
        
        GoSub RechercherArticle
        
        '-------- Modifier Article (Donn�es de base, CMS - CMS) --------
        Dim designation As String
        designation = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text
        valeur = InputBox("La designation du article " & article & " est : " & designation & Chr(13) _
        & "Ecrivez la nouvelle designation : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text = valeur

        GoSub Enregistrer
        
    End If
        
   If UserForm1.CheckBox2 = True Then 'Texte de commande
   
        GoSub RechercherArticle
        GoSub TexteDeCommande
        
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        Dim texteCommande As String
        texteCommande = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text
        valeur = InputBox("Le texte de commande du article " & article & " est : " & texteCommande & Chr(13) _
        & "Ecrivez le nouveau texte de commande : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = valeur
    
        GoSub Enregistrer
        
    End If

    If UserForm1.CheckBox3 = True Then 'Statut art. par div.

        GoSub RechercherArticle
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
        
    End If
        
    If UserForm1.CheckBox4 = True Then 'Type planification
    
        GoSub RechercherArticle
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim typePlanif As String
        typePlanif = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text
        valeur = InputBox("Le type de planification du article " & article & " est : " & typePlanif & Chr(13) _
        & "Ecrivez le nouveau type : ")
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = valeur
    
        'Il faut verifier si on a la bonne cl� pour le nouveau type de planification
        Dim cleCalcTailleLot As String
        cleCalcTailleLot = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text
        
        If (valeur = "VB" And cleCalcTailleLot = "") Then
            Select Case MsgBox("La cl� calc. taille lot n'est pas la bonnne pour le nouveau type 'VB'. Il faut la modifier !" _
            & " Voulez-vous modifier pour 'EX' ?", vbYesNo, "RPS")
                Case vbYes
                    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = "EX"
            End Select
        End If
        
        If (valeur = "ND" And cleCalcTailleLot = "EX") Then
            Select Case MsgBox("La cl� calc. taille lot n'est pas la bonnne pour le nouveau type 'ND'. Voulez-vous modifier pour '' ?", vbYesNo, "RPS")
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

    End If
    
    If UserForm1.CheckBox5 = True Then 'Point de commande
    
        GoSub RechercherArticle
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
         
    End If
    
    If UserForm1.CheckBox6 = True Then 'Valeur arrondie
    
        GoSub RechercherArticle
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
        
    End If

    If UserForm1.CheckBox7 = True Then 'D�lai livrai
    
        GoSub RechercherArticle
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

    End If
    
    If UserForm1.CheckBox8 = True Then 'Cl� calc. taille lot
    
        GoSub RechercherArticle
        GoSub MRP1
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim cleCalcTailleLot2 As String
        cleCalcTailleLot2 = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text
        valeur = InputBox("La cl� calc. taille lot du article " & article & " est : " & cleCalcTailleLot2 & Chr(13) _
        & "Ecrivez la nouvelle cl� : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = valeur
    
        typePlan = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text
    
        If (valeur = "" And typePlan = "VB") Then
            Select Case MsgBox("Le type planication n'est pas le bon pour la nouvelle cl� ''. Il faut le modifier !" _
            & " Voulez-vous modifier pour 'ND' ?", vbYesNo, "RPS")
                Case vbYes
                    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND"
                    typePlan = "ND"
            End Select
        End If
        
        If (valeur = "EX" And typePlan = "ND") Then
            Select Case MsgBox("Le type planification n'est pas le bon pour la nouvelle cl� 'EX'. Il faut le modifier !" _
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

    End If
    
    If UserForm1.CheckBox9 = True Then 'Num�ro pce. fabricant
    
        GoSub RechercherArticle
        GoSub Achats
        
        '-------- Modifier Article (Achats, CMS - CMS) --------
        Dim numFabricant As String
        numFabricant = session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text
        valeur = InputBox("Le num�ro pce fabricant du article " & article & " est : " & numFabricant & Chr(13) _
        & "Ecrivez le nouveau num�ro : ")
        session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text = valeur
    
        GoSub Enregistrer
    
    End If
    
    If UserForm1.CheckBox10 = True Then 'Emplacement
        
        GoSub RechercherArticle
        GoSub DonneesGenDivStockage
        
        '-------- Modifier Article (Donn�es g�n. div./stockage, CMS - CMS) --------
        Dim emplacement As String
        emplacement = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").Text
        valeur = InputBox("L'emplacement du article " & article & " est : " & emplacement & Chr(13) _
        & "Ecrivez le nouveau emplacement : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2701/txtMARD-LGPBE").Text = valeur
        
        session.findById("wnd[0]/tbar[1]/btn[18]").press 'Continuer
        
        '-------- Modifier Article (Gestion emplacements Masagin, CMS - CMS) --------
        session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text = valeur
    
        GoSub Enregistrer

    End If
    
    If UserForm1.CheckBox11 = True Then 'Grp Acheteur
    
        GoSub RechercherArticle
        GoSub Achats
        
        '-------- Modifier Article (Achats, CMS - CMS) --------
        Dim grpAcheteurs As String
        grpAcheteurs = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text
        valeur = InputBox("Le groupe acheteur du article " & article & " est : " & grpAcheteurs & Chr(13) _
        & "Ecrivez le nouveau groupe : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = valeur
        
        GoSub Enregistrer
    
    End If
    
    If UserForm1.CheckBox12 = True Then 'Gestionnaire
    
        GoSub RechercherArticle
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

    End If
    
    If UserForm1.CheckBox13 = True Then 'Cle Horizon
    
        GoSub RechercherArticle
        GoSub MRP1
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim cleHorizon As String
        cleHorizon = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text
        valeur = InputBox("La cl� horizon du article " & article & " est : " & cleHorizon & Chr(13) _
        & "Ecrivez la nouvelle cl� : ")
        session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text = valeur
        
        '[BUG]
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    End If
    
    If UserForm1.CheckBox14 = True Then 'Grp Marchandise
    
        GoSub RechercherArticle
    
        '-------- Modifier Article (Donn�es de base, CMS - CMS) --------
        Dim grpMarchandise As String
        grpMarchandise = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MATKL").Text
        valeur = InputBox("Le groupe merchandise du article " & article & " est : " & grpMarchandise & Chr(13) _
        & "Ecrivez le nouveau groupe : ")
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MATKL").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            MsgBox ("Vous avez annul� l'op�ration ! La session SAP sera ferm� !")
            Unload UserForm1
            fermetureSAP
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[1]/btn[18]").press 'Continuer
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARA-MATKL").Text = valeur
    
        GoSub Enregistrer
    
    End If
    
    If UserForm1.CheckBox15 = True Then 'Controle Dispo
    
        GoSub RechercherArticle
        GoSub MRP2
        
        '-------- Modifier article (MRP 2, CMS - CMS) --------
        Dim controleDispo As String
        controleDispo = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text
        valeur = InputBox("Le controle disponibil. du article " & article & " est : " & controleDispo & Chr(13) _
        & "Ecrivez le nouveau controle : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = valeur
    
        GoSub Enregistrer

    End If
    
    If UserForm1.CheckBox16 = True Then 'Type magasin pour SM
    
        GoSub RechercherArticle
        GoSub GestionEmplacementsMagasin
        
        '-------- Modifier Article (Gestion emplacements Masagin, CMS - CMS) --------
        Dim typeMagSM As String
        typeMagSM = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text
        valeur = InputBox("Le type magasin pour SM du article " & article & " est : " & typeMagSM & Chr(13) _
        & "Ecrivez le nouveau type : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text = valeur
        
        GoSub Enregistrer

    End If
    
    If UserForm1.CheckBox17 = True Then 'Type magasin EM
    
        GoSub RechercherArticle
        GoSub GestionEmplacementsMagasin
        
        '-------- Modifier Article (Gestion emplacements Masagin, CMS - CMS) --------
        Dim typeMagEM As String
        typeMagEM = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text
        valeur = InputBox("Le type magasin EM du article " & article & " est : " & typeMagEM & Chr(13) _
        & "Ecrivez le nouveau type : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text = valeur
        
        GoSub Enregistrer
    
    End If
    
'    Unload UserForm1
    compteur = compteur + 1
    
Next i

Unload UserForm1

'Sauvegarder
Workbooks(fichier).Save

MsgBox ("Vous avez modifi� " & compteur & " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

Exit Sub

RechercherArticle:
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier Article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    Return

Enregistrer:
    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        MsgBox ("Vous avez annul� l'op�ration ! La session SAP sera ferm�e !")
        Unload UserForm1
        fermetureSAP
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne � l'ecran initial)
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    Return

'[BUG]
Enregistrer2:
    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        MsgBox ("Vous avez annul� l'op�ration ! La session SAP sera ferm�e !")
        Unload UserForm1
        fermetureSAP
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne � l'ecran initial)
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    Return

Achats:
    '-------- Modifier Article (Donn�es de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return
    
TexteDeCommande:
    '-------- Modifier Article (Donn�es de base, CMS - CMS) --------
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
    'Si on a "ND", il y a une �tape de plus, une message est affich�e
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
    
    '-------- Modifier Article (Donn�es g�n. div./stockage, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return

End Sub
