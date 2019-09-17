Attribute VB_Name = "modifierArticles"
Option Explicit

'Modifier tous les articles dans la liste, l'utilisateur doit choisir quelle(s) modification(s) à faire
'et la(es) nouvelle(s) valeur(s) à debut de la procedure. Aprés une boucle modifie tous les articles.
'Modifier les articles pour Nantes ou Saint-Nazaire, la liste doit avoir des articles d'un site ou de
'l'autre, ne pas mélanger

Sub modifierArticles_SAP()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                    'Modifier les article
                    
Dim article As String, fichier As String, i As Integer, fin As String, ref As String
Dim ws As Worksheet

fichier = ThisWorkbook.Name
Set ws = Workbooks(fichier).Worksheets("PREPA SAP")

fin = ws.Cells(Rows.Count, 2).End(xlUp).Row

'Configurer parametrages
ref = 4
article = ws.Range("B" & ref).Value

'Modifier l'article pour le site à Nantes ou à Saint Nazaire
Dim division As String, magasin As String, numeroMagasin As String, typeMagasin As String, valeur As String
Dim typePlan As String, compteur As Integer

division = ws.Range("J" & ref).Value 'NTF ou (NZF)
magasin = ws.Range("K" & ref).Value 'NENM ou (Z62M)
numeroMagasin = ws.Range("L" & ref).Value 'N18 ou (Z18)
typeMagasin = ws.Range("M" & ref).Value 'NEN ou (Z62)

'Sauvegarder le niveaux de organisation et la selection des vues
'Call enregistrerOrganisationEtVues(article, ref, division, magasin, numeroMagasin, typeMagasin)

compteur = 0

Load UserForm1 'créer l'UserForm, mais pas l'afficher
UserForm1.TextBox1 = article
UserForm1.TextBox2 = ws.Range("B" & fin).Value

MsgBox ("Choisissez le(s) modification(s) à faire pour tous les articles!")
UserForm1.Show

'Vérifier si l'utilisateur fait au moins un choix
If verifierOptions(UserForm1.Controls) = False Then
    MsgBox ("Vous n'avez pas choisi faire des modifications ! La session SAP sera fermée !")
    Unload UserForm1
    fermetureSAP
    Exit Sub
End If

'Demander des entrees
Const qteOptions As Integer = 17
Dim entrees(1 To qteOptions) As String, variable(1 To qteOptions) As String

variable(1) = "de la designation"
variable(2) = "du texte de commande"
variable(3) = "du statut art. par div."
variable(4) = "du type de planification"
variable(5) = "du point de commande"
variable(6) = "de la valeur arrondie"
variable(7) = "du delai livrais"
variable(8) = "de la clé calc. taille lot"
variable(9) = "du numéro pce. fabricant"
variable(10) = "de l'emplacement"
variable(11) = "du groupe acheteurs"
variable(12) = "du gestionnaire"
variable(13) = "de la clé horizon"
variable(14) = "du groupe merchandise"
variable(15) = "du controle dispo"
variable(16) = "du type magasin pour SM"
variable(17) = "du type magasin EM"

For i = 1 To 17
    
    If UserForm1.Controls("CheckBox" & i).Value = True Then
        entrees(i) = InputBox("Ecrivez la nouvelle valeur " + variable(i) + " pour les articles :")
        verifierEntree (entrees(i))
    End If

Next i

'Boucle pour faire des modifications
For i = 4 To fin

    article = ws.Range("B" & i).Value
    
    'Vérifier quelle(s) modification(s) à faire
    If UserForm1.CheckBox1 = True Then 'Designation
        
        GoSub RechercherArticle
        
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        valeur = entrees(1)
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text = valeur

        GoSub Enregistrer
        
    End If
        
   If UserForm1.CheckBox2 = True Then 'Texte de commande
   
        GoSub RechercherArticle
        GoSub TexteDeCommande
        
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        valeur = entrees(2)
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = valeur
    
        GoSub Enregistrer
        
    End If

    If UserForm1.CheckBox3 = True Then 'Statut art. par div.

        GoSub RechercherArticle
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        valeur = entrees(3)
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text = valeur
    
        'Etape suplementaire
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
        valeur = entrees(4)
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
        
        'Etape suplementaire
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
        valeur = entrees(5)
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = valeur
        
        'Etape suplementaire
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
        valeur = entrees(6)
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text = valeur
        
        'Etape suplementaire
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If
        
    End If

    If UserForm1.CheckBox7 = True Then 'Délai livrai
    
        GoSub RechercherArticle
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        valeur = entrees(7)
        session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text = valeur
        
        'Etape suplementaire
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    End If
    
    If UserForm1.CheckBox8 = True Then 'Clé calc. taille lot
    
        GoSub RechercherArticle
        GoSub MRP1
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        valeur = entrees(8)
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
        
        'Etape suplementaire
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    End If
    
    If UserForm1.CheckBox9 = True Then 'Numéro pce. fabricant
    
        GoSub RechercherArticle
        GoSub Achats
        
        '-------- Modifier Article (Achats, CMS - CMS) --------
        valeur = entrees(9)
        session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text = valeur
    
        GoSub Enregistrer
    
    End If
    
    If UserForm1.CheckBox10 = True Then 'Emplacement
        
        GoSub RechercherArticle
        GoSub DonneesGenDivStockage
        
        '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
        valeur = entrees(10)
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
        valeur = entrees(11)
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = valeur
        
        GoSub Enregistrer
    
    End If
    
    If UserForm1.CheckBox12 = True Then 'Gestionnaire
    
        GoSub RechercherArticle
        GoSub MRP1
        
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        valeur = entrees(12)
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text = valeur
        
        'Etape suplementaire
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
        valeur = entrees(13)
        session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text = valeur
        
        'Etape suplementaire
        If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
            GoSub Enregistrer2
        Else
            GoSub Enregistrer
        End If

    End If
    
    If UserForm1.CheckBox14 = True Then 'Grp Marchandise
    
        GoSub RechercherArticle
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        valeur = entrees(14)
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MATKL").Text = valeur
        session.findById("wnd[0]/tbar[1]/btn[18]").press 'Continuer
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        valeur = entrees(14)
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARA-MATKL").Text = valeur
    
        GoSub Enregistrer
    
    End If
    
    If UserForm1.CheckBox15 = True Then 'Controle Dispo
    
        GoSub RechercherArticle
        GoSub MRP2
        
        '-------- Modifier article (MRP 2, CMS - CMS) --------
        valeur = entrees(15)
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = valeur
    
        GoSub Enregistrer

    End If
    
    If UserForm1.CheckBox16 = True Then 'Type magasin pour SM
    
        GoSub RechercherArticle
        GoSub GestionEmplacementsMagasin
        
        '-------- Modifier Article (Gestion emplacements Masagin, CMS - CMS) --------
        valeur = entrees(16)
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text = valeur
        
        GoSub Enregistrer

    End If
    
    If UserForm1.CheckBox17 = True Then 'Type magasin EM
    
        GoSub RechercherArticle
        GoSub GestionEmplacementsMagasin
        
        '-------- Modifier Article (Gestion emplacements Masagin, CMS - CMS) --------
        valeur = entrees(17)
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text = valeur
        
        GoSub Enregistrer
    
    End If
    
    compteur = compteur + 1
    
Next i

Unload UserForm1

'Sauvegarder
Workbooks(fichier).Save

MsgBox ("Vous avez modifié " & compteur & " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

Exit Sub

'Démarrer l'opération de modification
RechercherArticle:
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier Article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    
    Return

'Enregistrer les modifications
Enregistrer:
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    Return

'Etape suplementaire
Enregistrer2:
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    Return

'Avancer les fenêtres
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

Function verifierOptions(formulaireControls As Controls) As Boolean

Dim ctrl As Control
verifierOptions = False

For Each ctrl In formulaireControls
    If TypeName(ctrl) = "CheckBox" Then
        If ctrl.Value = True Then
            verifierOptions = True
        End If
    End If
Next ctrl

End Function

Sub verifierEntree(var As String)

If StrPtr(var) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
    MsgBox ("Vous avez annulé l'opération ! La session SAP sera fermée !")
    Unload UserForm1
    fermetureSAP
    End
End If

End Sub

Sub enregistrerOrganisationEtVues(article As String, ref As String, division As String, magasin As String, numeroMagasin As String, typeMagasin As String)

'-------- Barre de recherche --------
session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
session.findById("wnd[0]").sendVKey 0

'-------- Modifier Article (Ecran initial) --------
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article

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
session.findById("wnd[1]/tbar[0]/btn[5]").press 'Sélection des vues

'Selection des vues
session.findById("wnd[1]/tbar[0]/btn[19]").press 'Demarquer tout

If division = "NTF" Then 'Nantes

    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Données de base
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).Selected = True 'Achats
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).Selected = True 'Texte de commande
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).Selected = True 'MRP 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'MRP 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Données gén. div./stockage
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'Gestion emplacements magasin
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'Comptabilité

ElseIf (division = "NZF") Then

    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Données de base
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).Selected = True 'Achats
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).Selected = True 'Texte de commande
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).Selected = True 'MRP 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).Selected = True 'MRP 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'Données gén. div./stockage
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Gestion emplacements magasin
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'Comptabilité

End If

session.findById("wnd[1]/tbar[0]/btn[14]").press 'Sauvegarder comme paramétrage
session.findById("wnd[1]/tbar[0]/btn[6]").press 'Niveaux organisation
session.findById("wnd[1]/tbar[0]/btn[13]").press 'Sauvegarder comme paramétrage
session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour (F3)
session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour (F3)


End Sub
