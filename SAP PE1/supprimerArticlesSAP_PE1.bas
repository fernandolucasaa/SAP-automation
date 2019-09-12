Attribute VB_Name = "supprimerArticlesSAP_PE1"
Option Explicit

Sub supprimerArticles_SAPPE1()

Dim fichier As String, article As String
Dim premier As Integer, dernier As Integer, i As Integer, compteur As Integer
Dim ws As Worksheet

fichier = ThisWorkbook.Name
Set ws = Windows(fichier).ActiveSheet

Workbooks(fichier).Activate
premier = 2
dernier = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
compteur = 0

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                'Supprimer des articles

'Boucle pour la modification des articles
For i = premier To dernier
    
    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm06"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Position témoin suppresion article : écran de sélection --------
    Dim division As String, emplStockage As String
    
    article = ws.Range("A" & i).Value
    division = ws.Range("B" & i).Value
    emplStockage = ws.Range("C" & i).Value
    
    session.findById("wnd[0]/usr/ctxtRM03G-MATNR").Text = article
    session.findById("wnd[0]/usr/ctxtRM03G-WERKS").Text = division
    session.findById("wnd[0]/usr/ctxtRM03G-LGORT").Text = emplStockage
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Position témoin suppresion article : écran de données --------
    session.findById("wnd[0]/usr/chkRM03G-LVOLG").Selected = True 'Emplacement stockage
       
    session.findById("wnd[0]/tbar[0]/btn[11]").press 'Sauvegarder (on retourne à l'ecran initial)
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
       
    compteur = compteur + 1

Next i

MsgBox ("Vous avez supprimé " & compteur & " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If
    
End Sub
