Attribute VB_Name = "recupererDataSAP_PE1"
Option Explicit

Sub recupererData_SAPPE1()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                'Recuperer DATA
Dim fichier As String, article As String
Dim premier As Integer, dernier As Integer, i As Integer, compteur As Integer

fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
premier = 2
dernier = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
compteur = 0

For i = premier To dernier

    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm03"
    session.findById("wnd[0]").sendVKey 0

    '-------- Afficher article (Ecran initial) --------
    Windows(fichier).Activate
    article = ActiveSheet.Range("A" & i).Value
    
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article
    session.findById("wnd[0]/tbar[1]/btn[6]").press 'Niveaux organisation
    
    'Nvx organisationnels
    Dim division As String, emplStockage As String
    
    Windows(fichier).Activate
    division = ActiveSheet.Range("B" & i).Value 'ME01
    emplStockage = ActiveSheet.Range("C" & i).Value 'OM05
    
    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = division
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = emplStockage
    session.findById("wnd[1]/tbar[0]/btn[5]").press 'Selection des vues

    'Selection des vues (ouvri de mani�ere automatique)
    session.findById("wnd[1]/tbar[0]/btn[0]").press 'Effacer la s�lection
'    session.findById("wnd[1]/tbar[0]/btn[19]").press 'Effacer la s�lection
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Donn�es de base 1
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).Selected = True 'Achats
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).Selected = True 'Texte de commande d'achat
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'Planification des besoins 1
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Planification des besoins 2
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'Planification des besoins 3
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'Planification des besoins 4
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).Selected = True 'Donn�es g�n. divs./stockage 1
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).Selected = True 'Donn�es g�n. divs./stockage 2
'    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).Selected = True 'Comptabilit� 1
'    session.findById("wnd[1]/tbar[0]/btn[0]").press 'Suite

    '-------- Afficher article (Donn�es de base 1) --------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Afficher article (Achats) --------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Afficher article (Texte commande de achat) --------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Afficher article (Planif. des besions 1) -------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Afficher article (Planif. des besions 2) -------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Afficher article (Planif. des besions 3) -------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Afficher article (Planif. des besions 4) -------
    Dim pointCommande As String, qteReapprov As String
    
    pointCommande = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/txtMARD-LMINB").Text
    qteReapprov = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB6:SAPLMGD1:2498/txtMARD-LBSTF").Text
        
    Windows(fichier).Activate
    ActiveSheet.Range("D" & i).Value = pointCommande
    ActiveSheet.Range("E" & i).Value = qteReapprov
        
    session.findById("wnd[0]").sendVKey 0

    '-------- Afficher article (Donn.div./stockage 1) -------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Afficher article (Donn.div./stockage 2) -------
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Afficher article (Comptabilit� 1) -------
    session.findById("wnd[0]").sendVKey 0
    
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press 'Quitter l'affichage de l'article

    compteur = compteur + 1
    
Next i

MsgBox ("Vous avez consult� " & compteur & " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

End Sub
