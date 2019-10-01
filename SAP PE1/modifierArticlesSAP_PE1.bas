Attribute VB_Name = "modifierArticlesSAP_PE1"
Option Explicit

Sub modifierArticles_SAPPE1()

Dim fichier As String, article As String
Dim premier As Integer, dernier As Integer, i As Integer, compteur As Integer

fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
'premier = Selection.Row
'dernier = premier + Selection.Rows.Count - 1
premier = 2
dernier = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
compteur = 0

Load UserForm1 'créer l'UserForm, mais pas l'afficher
UserForm1.TextBox1 = ActiveSheet.Range("A" & premier).Value
UserForm1.TextBox2 = ActiveSheet.Range("A" & dernier).Value
MsgBox ("Choisissez la modification à faire pour tous les articles selectionés !")
UserForm1.Show

Dim valeur1, valeur2 As String

If UserForm1.OptionButton1.Value = True Then 'Type de planification et statut art. par div.

    valeur1 = InputBox("Ecrivez le nouveau type de planification :")
    verifierEntree (valeur1)
    valeur2 = InputBox("Ecrivez le nouveau statut art. par div. :")
    verifierEntree (valeur2)

ElseIf UserForm1.OptionButton2.Value = True Then 'Point de commande

    valeur1 = InputBox("Ecrivez le nouveau point de commande :")
    verifierEntree (valeur1)

ElseIf UserForm1.OptionButton3.Value = True Then 'Taille de lot fixe

    valeur1 = InputBox("Ecrivez la nouvelle taille de lot fixe :")
    verifierEntree (valeur1)

ElseIf UserForm1.OptionButton4.Value = True Then 'Emplacement

    valeur1 = InputBox("Ecrivez le nouvel emplacement :")
    verifierEntree (valeur1)

ElseIf UserForm1.OptionButton5.Value = True Then 'Texte de commane

    valeur1 = InputBox("Ecrivez le nouveau texte de commande :")
    verifierEntree (valeur1)

End If

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                'Modifier des articles

'Boucle pour la modification des articles
For i = premier To dernier

    Workbooks(fichier).Activate
    article = ActiveSheet.Range("A" & i).Value
    
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm02"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier Article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
    session.findById("wnd[0]").sendVKey 0
    
    'Selection des vues (Pour division NZ01)
    session.findById("wnd[1]/tbar[0]/btn[19]").press 'Effacer la sélection
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Données de base 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).Selected = True 'Achats
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).Selected = True 'Texte de commande d'achat
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'Planification des besoins 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Planification des besoins 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'Planification des besoins 3
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'Planification des besoins 4
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).Selected = True 'Données gén. divs./stockage 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).Selected = True 'Données gén. divs./stockage 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).Selected = True 'Comptabilité 1
    session.findById("wnd[1]/tbar[0]/btn[0]").press 'Enter

    If UserForm1.OptionButton1.Value = True Then 'Type de planification et statut art. par div.
    
        '-------- Modifier article (Données de base 1) --------
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MSTAE").Text = valeur2 'statut art. par div.
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MSTAE").SetFocus
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP01/ssubTABFRA1:SAPLMGMM:2004/subSUB4:SAPLMGD1:2001/ctxtMARA-MSTAE").caretPosition = 2
        session.findById("wnd[0]").sendVKey 0
        
        '-------- Modifier article (Achats) --------
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-MMSTA").Text = valeur2 'statut art. par div.
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1001/txtMAKT-MAKTX").caretPosition = 10
        session.findById("wnd[0]").sendVKey 0
        
        '-------- Modifier article (Texte commande de achat) --------
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX").SetFocus
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX").caretPosition = 22
        session.findById("wnd[0]").sendVKey 0
        
        '-------- Modifier article (Planif. des besions 1) -------
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").Text = valeur1 'type planification
    
        'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
        
    ElseIf UserForm1.OptionButton2.Value = True Then 'Point de commande
        
        GoSub PlanifDesBesoins1
        
        '-------- Modifier article (Planif. des besoins 1) -------
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/txtMARC-MINBE").Text = valeur1
        
        'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm1.OptionButton3.Value = True Then 'Taille de lot fixe
        
        GoSub PlanifDesBesoins1
        
        '-------- Modifier article (Planif. des besoins 1) -------
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/txtMARC-BSTFE").Text = valeur1
    
        'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
        
    ElseIf UserForm1.OptionButton4.Value = True Then 'Emplacement
        
        GoSub DonnDivStockage1
        
        '-------- Modifier article (Donn.div./stockage 1) -------
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLZMGD1:2701/txtMARD-LGPBE").Text = valeur1
    
        'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
    
    ElseIf UserForm1.OptionButton5.Value = True Then 'Texte de commande
    
        '-------- Modifier article (Données de base 1) --------
        session.findById("wnd[0]").sendVKey 0
        
        '-------- Modifier article (Achats) --------
        session.findById("wnd[0]").sendVKey 0
        
        '-------- Modifier article (Texte commande de achat) --------
        session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").Text = valeur1
    
        'Sauvegarder
        session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
        
    End If
    
    compteur = compteur + 1

Next i

'Modification finie
Unload UserForm1

MsgBox ("Vous avez modifié " & compteur & " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

Exit Sub

PlanifDesBesoins1:

    '-------- Modifier article (Données de base 1) --------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier article (Achats) --------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier article (Texte commande de achat) --------
    session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB1:SAPLMGD1:1002/txtMAKT-MAKTX").SetFocus
    session.findById("wnd[0]").sendVKey 0
    
    Return

DonnDivStockage1:

    GoSub PlanifDesBesoins1
    
    '-------- Modifier article (Planif. des besions 1) -------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier article (Planif. des besions 2) -------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier article (Planif. des besions 3) -------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Modifier article (Planif. des besions 4) -------
    session.findById("wnd[0]").sendVKey 0

    Return
    
End Sub

Sub verifierEntree(v As String)
    
If StrPtr(v) = 0 Then 'Cliquer sur 'Annuler' ou fermer la fenetre
    MsgBox ("Vous avez annulé l'opération !")
    Unload UserForm1
    End
End If

End Sub
