Attribute VB_Name = "suppresionArticlesSAP_PE1"
Option Explicit

Sub suppresionArticles_SAPPE1()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                'Suppression des articles
Dim fichier As String, article As String
Dim premier As Integer, dernier As Integer, i As Integer, compteur As Integer

fichier = ThisWorkbook.Name

Workbooks(fichier).Activate
premier = Selection.Row
dernier = premier + Selection.Rows.Count - 1
compteur = 0

'Boucle pour supprimer des articles
For i = premier To dernier

    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm06"
    session.findById("wnd[0]").sendVKey 0
    
    '-------- Position t�moin suppresion article : �cran de sel�ction --------
    Dim division As String, emplStockage As String
    
    session.findById("wnd[0]/usr/ctxtRM03G-MATNR").Text = article
    session.findById("wnd[0]/usr/ctxtRM03G-WERKS").Text = division
    session.findById("wnd[0]/usr/ctxtRM03G-LGORT").Text = emplStockage

Next i

MsgBox ("Vous avez supprim� " & compteur & " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

End Sub
