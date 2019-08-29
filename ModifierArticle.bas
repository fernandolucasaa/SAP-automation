Attribute VB_Name = "Module2"
Option Explicit

Sub modifierArticle()

'_________________________________________________________________________________________________'
                    'Logon SAP
'Variables
Dim SapGui, Applic, Connection, session, WSHShell
Dim identifiant As String, motDePasse As String, langue As String

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
                   'Modifier Article
Dim article As String, fichier As String, i As Integer

i = Selection.Row
fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
article = ActiveSheet.Range("B" & i).Value

Load UserForm1 'creer l'UserForm, mais pas l'afficher
UserForm1.TextBox1 = article

Workbooks(fichier).Activate
UserForm1.Show

'-------- Barre de recherche --------
session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
session.findById("wnd[0]").sendVKey 0

'-------- Modifier Article (Ecran initial) --------
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
'session.findById("wnd[0]/tbar[0]/btn[0]").press 'retirer le commentaire si on ne veut pas configurer le niveau de org.

'Modifier l'article pour le site à Nantes ou à Saint Nazaire
Dim division As String, magasin As String, numeroMagasin As String, typeMagasin As String, valeur As String

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
session.findById("wnd[0]/tbar[0]/btn[0]").press

'Il faut selectionner les vues aussi ?

'[BUG]
Dim typePlan As String

If UserForm1.OptionButton1 = True Then 'Designation
    
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    Dim designation As String
    designation = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text
    valeur = InputBox("La designation du article " & article & " est : " & designation & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text = valeur
    
    GoSub Enregistrer
    
ElseIf UserForm1.OptionButton2 = True Then 'Texte de commande

    GoSub TexteDeCommande
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    Dim texteCommande As String
    texteCommande = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text
    valeur = InputBox("Le texte de commande du article " & article & " est : " & texteCommande & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = valeur

    GoSub Enregistrer

ElseIf UserForm1.OptionButton3 = True Then 'Statut art. par div.

    GoSub MRP1
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim statutArt As String
    statutArt = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text
    valeur = InputBox("Le statut art. par div. du article " & article & " est : " & statutArt & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text = valeur

    GoSub Enregistrer

ElseIf UserForm1.OptionButton4 = True Then 'Type planification
    
    GoSub MRP1
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim typePlanif As String
    typePlanif = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text
    valeur = InputBox("Le type de planification du article " & article & " est : " & typePlanif & Chr(13) & "Ecrivez le nouveau valeur : ")
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
    
    GoSub Enregistrer
    
    '[BUG]
    'Une fois qu'on change le VB pour ND, quand on veut enregistrer, il y a une message differente,
    'donc, quand on fait le retour on ouvre une fenetre pour quitter
    If (valeur = "ND") Then
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    End If

ElseIf UserForm1.OptionButton5 = True Then 'Point de commande

    GoSub MRP1
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim ptCommande As String
    ptCommande = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text
    valeur = InputBox("Le point de commande du article " & article & " est : " & ptCommande & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = valeur
    
    '[BUG]
    typePlan = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text

    GoSub Enregistrer
    
    '[BUG]
    'Même avertissement indiqué dessus, une fois que on est dans MRP1 et on a VB
    'une fenetre est ouverte apres essayer de quitter l'opération
    If (typePlan = "ND") Then
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    End If
         
ElseIf UserForm1.OptionButton6 = True Then 'Valeur arrondie

    GoSub MRP1
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim valeurArrondie As String
    valeurArrondie = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text
    valeur = InputBox("Le valeur arrondie du article " & article & " est : " & valeurArrondie & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text = valeur

    '[BUG]
    typePlan = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text
    
    GoSub Enregistrer
    
    '[BUG]
    If (typePlan = "ND") Then
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    End If

ElseIf UserForm1.OptionButton7 = True Then 'Délai livrai

    GoSub MRP1
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim delaiLivrai As String
    delaiLivrai = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text
    valeur = InputBox("Le delai livrais du article " & article & " est : " & delaiLivrai & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text = valeur

    '[BUG]
    typePlan = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text

    GoSub Enregistrer
    
    '[BUG]
    If (typePlan = "ND") Then
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    End If

ElseIf UserForm1.OptionButton8 = True Then 'Clé calc. taille lot

    GoSub MRP1

    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim cleCalcTailleLot2 As String
    cleCalcTailleLot2 = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text
    valeur = InputBox("La clé calc. taille lot du article " & article & " est : " & cleCalcTailleLot2 & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = valeur

    '[BUG]
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
    
    GoSub Enregistrer
    
    '[BUG]
    If (typePlan = "ND") Then
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    End If

ElseIf UserForm1.OptionButton9 = True Then 'Numéro pce. fabricant

    GoSub Achats
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    Dim numFabricant As String
    numFabricant = session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text
    valeur = InputBox("Le numéro pce fabricant du article " & article & " est : " & numFabricant & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text = valeur

    GoSub Enregistrer
    
ElseIf UserForm1.OptionButton10 = True Then 'Emplacement
    
    GoSub GestionEmplacementsMagasin
    
    '-------- Gestion emplacements Masagin (CMS - CMS) --------
    Dim emplacement As String
    emplacement = session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text
    valeur = InputBox("L'emplacement du article " & article & " est : " & emplacement & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text = valeur

    GoSub Enregistrer

ElseIf UserForm1.OptionButton11 = True Then 'Grp Acheteur (Gestionnaire)

    GoTo Achats
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    Dim grpAcheteurs As String
    grpAcheteurs = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text
    valeur = InputBox("Le groupe acheteur du article " & article & " est : " & grpAcheteurs & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").Text = grpAcheteurs

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If

    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Continuer

    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press

    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim gestionnaire As String
    gestionnaire = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text
    valeur = InputBox("Le gestionnaire du article " & article & " est : " & gestionnaire & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").Text = gestionnaire

    GoSub Enregistrer

ElseIf UserForm1.OptionButton12 = True Then 'Cle Horizon

    GoSub MRP1

    '-------- Modifier Article (MRP1, CMS - CMS) --------
    Dim cleHorizon As String
    cleHorizon = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text
    valeur = InputBox("La clé horizon du article " & article & " est : " & cleHorizon & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/ctxtMARC-FHORI").Text = cleHorizon

    GoSub Enregistrer

ElseIf UserForm1.OptionButton13 = True Then 'Grp Marchandise

    '-------- Modifier Article (Données de base, CMS - CMS) --------
    Dim grpMarchandise As String
    grpMarchandise = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MATKL").Text
    valeur = InputBox("Le groupe merchandise du article " & article & " est : " & grpMarchandise & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2001/ctxtMARA-MATKL").Text = grpMarchandise

    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If

    session.findById("wnd[0]/tbar[1]/btn[18]").press 'Continuer

    '-------- Modifier Article (Achats, CMS - CMS) --------
    grpMarchandise = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARA-MATKL").Text
    valeur = InputBox("Le groupe merchandise du article " & article & " est : " & grpMarchandise & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2301/ctxtMARA-MATKL").Text = grpMarchandise

    GoSub Enregistrer
    
ElseIf UserForm1.OptionButton14 = True Then 'Controle Dispo

    GoSub MRP2
    
    '-------- Modifier article (MRP 2, CMS - CMS) --------
    Dim controleDispo As String
    controleDispo = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text
    valeur = InputBox("Le controle disponibil. du article " & article & " est : " & controleDispo & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2493/ctxtMARC-MTVFP").Text = controleDispo

    GoSub Enregistrer

ElseIf UserForm1.OptionButton15 = True Then 'Type magasin pour SM

    GoSub GestionEmplacementsMagasin
    
    '-------- Gestion emplacements Masagin (CMS - CMS) --------
    Dim typeMagSM As String
    typeMagSM = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text
    valeur = InputBox("Le type magasin pour SM du article " & article & " est : " & typeMagSM & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZA").Text = valeur
    
    GoSub Enregistrer

ElseIf UserForm1.OptionButton16 = True Then 'Type magasin EM

    GoSub GestionEmplacementsMagasin
    
    '-------- Gestion emplacements Masagin (CMS - CMS) --------
    Dim typeMagEM As String
    typeMagEM = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text
    valeur = InputBox("Le type magasin EM du article " & article & " est : " & typeMagEM & Chr(13) & "Ecrivez le nouveau valeur : ")
    session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2733/ctxtMLGN-LTKZE").Text = valeur
    
    GoSub Enregistrer
    
End If

Unload UserForm1

'Sauvegarder
Workbooks(fichier).Save

'Fermeture de la connexion
If MsgBox("La modification des articles est fini. Voulez-vous fermer votre session SAP ?", vbYesNo, "RPS") = vbYes Then
    session.findById("wnd[0]").Close
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
End If

Exit Sub

Enregistrer:
    If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
        session.findById("wnd[0]").Close
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        Exit Sub
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    Return  'Retour

Achats:
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return  'Retour
    
TexteDeCommande:
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return  'Retour

MRP1:
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return  'Retour

MRP2:
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return  'Retour
     
GestionEmplacementsMagasin:
    '-------- Modifier Article (Données de base, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Achats, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Texte de commande, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (MRP1, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    '[BUG]
    'Si on a "ND", avant executer une action, il y a une message
    If (session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = "ND") Then
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    End If
    
    '-------- Modifier Article (MRP2, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
    session.findById("wnd[0]/tbar[1]/btn[18]").press
    
    Return  'Retour
    
End Sub
