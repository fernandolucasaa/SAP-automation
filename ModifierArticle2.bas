Attribute VB_Name = "Module3"
Option Explicit

Sub modifierArticles2()

'_________________________________________________________________________________________________'
                    'Logon SAP
'Variables
Dim SapGui, Applic, Connection, session, WSHShell
Dim identifiant As String, motDePasse As String, langue As String

'identifiant = "ng2b609"
'motDePasse = "Dr210591"
'identifiant = "ng2b23d"
'motDePasse = "RPS08201"

identifiant = InputBox("Ecrivez votre identifiant de l'utilisateur", "RPS")
If StrPtr(identifiant) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
    Exit Sub
End If

motDePasse = InputBox("Ecrivez votre mot de passe", "RPS")
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

Dim i As Integer, premier As Integer, dernier As Integer, article As String
Dim fichier As String, fin As String

fichier = ThisWorkbook.Name
fin = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row

'premier = Selection.Row
'dernier = Selection.Rows.Count + premier - 1

Load UserForm2 'creer l'UserForm, mais par l'afficher

'For i = premier To dernier
For i = 4 To fin

    Workbooks(fichier).Activate
    article = ActiveSheet.Range("B" & i).Value
    UserForm2.TextBox1 = article
    UserForm2.Show
    
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier Article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
    
    'Modifier l'article pour le site à Nantes ou à Saint Nazaire
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
    session.findById("wnd[0]/tbar[0]/btn[0]").press

    'Il faut selectionner les vues aussi ?
    
    Dim valeur As String
    
    If UserForm2.OptionButton1 = True Then 'Designation
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        Dim designation As String
        designation = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text
        valeur = InputBox("La designation du article " & article & " est : " & designation & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:8001/tblSAPLMGD1TC_KTXT/txtSKTEXT-MAKTX[1,0]").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    
    ElseIf UserForm2.OptionButton2 = True Then 'Texte de commande
    
         '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        Dim texteCommande As String
        texteCommande = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text
        valeur = InputBox("Le texte de commande du article " & article & " est : " & texteCommande & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm2.OptionButton3 = True Then 'Statut art. par div.
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim statutArt As String
        statutArt = session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text
        valeur = InputBox("Le statut art. par div. du article " & article & " est : " & statutArt & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB2:SAPLMGD1:2481/ctxtMARC-MMSTA").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm2.OptionButton4 = True Then 'Type planification
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim typePlanif As String
        typePlanif = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text
        valeur = InputBox("Le type de planification du article " & article & " est : " & typePlanif & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
        If (valeur = "ND") Then
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        End If
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm2.OptionButton5 = True Then 'Point de commande
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim ptCommande As String
        ptCommande = session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text
        valeur = InputBox("Le point de commande du article " & article & " est : " & ptCommande & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm2.OptionButton6 = True Then 'Valeur arrondie
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim valeurArrondie As String
        valeurArrondie = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text
        valeur = InputBox("Le valeur arrondie du article " & article & " est : " & valeurArrondie & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/txtMARC-BSTRF").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm2.OptionButton7 = True Then 'Délai livrai
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim delaiLivrai As String
        delaiLivrai = session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text
        valeur = InputBox("Le delai livrais du article " & article & " est : " & delaiLivrai & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB7:SAPLMGD1:2485/txtMARC-PLIFZ").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm2.OptionButton8 = True Then 'Clé calc. taille lot
    
         '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        Dim cleCalcTailleLot As String
        cleCalcTailleLot = session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text
        valeur = InputBox("La clé calc. taille lot du article " & article & " est : " & cleCalcTailleLot & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm2.OptionButton9 = True Then 'Numéro pce. fabricant
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        Dim numFabricant As String
        numFabricant = session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text
        valeur = InputBox("Le numéro pce fabricant du article " & article & " est : " & numFabricant & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB11:SAPLMGD1:2312/txtMARA-MFRPN").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm2.OptionButton10 = True Then 'Emplacement
    
        '-------- Modifier Article (Données de base, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Achats, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Texte de commande, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (MRP1, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
        'session.findById("wnd[0]/tbar[1]/btn[18]").press 'pas necessaire ?
    
        '-------- Modifier Article (MRP2, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
        'session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Modifier Article (Données gén. div./stockage, CMS - CMS) --------
        session.findById("wnd[0]/tbar[1]/btn[18]").press
    
        '-------- Gestion emplacements Masagin (CMS - CMS) --------
        Dim emplacement As String
        emplacement = session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text
        valeur = InputBox("L'emplacement du article " & article & " est : " & emplacement & Chr(13) & "Ecrivez le nouveau valeur : ")
        session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text = valeur
    
        If StrPtr(valeur) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
            session.findById("wnd[0]").Close
            session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
            Exit Sub
        End If
    
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    End If
    
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    'session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
Next i

Unload UserForm2

'Sauvegarder
Workbooks(fichier).Save

'Fermeture de la connexion
If MsgBox("La modification des articles est fini. Voulez-vous fermer votre session SAP ?", vbYesNo, "RPS") = vbYes Then
    session.findById("wnd[0]").Close
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
End If

End Sub


