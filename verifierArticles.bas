Attribute VB_Name = "verifierArticles"
Option Explicit

'V�rifier si les articles sont vraiment bien cr��es, c'est-�-dire que une fois qu'on essaye de cr�er
'une article d�j� cr�� une message d'erreur appara�t
'V�rifier des articles pour Nantes et Saint-Nazaire

Sub verifierArticles_SAP()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                        'V�rifier des articles
Dim article As String, modele As String, fichier As String, i As Integer, fin As String
Dim articlesIncomplets As String, articlesDejaCrees As String, compteur As Integer, cpt As Integer

fichier = ThisWorkbook.Name
fin = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row
compteur = 0 'articles pas finis
cpt = 0 'articles finis

For i = 4 To fin

    Workbooks(fichier).Activate
    article = ActiveSheet.Range("B" & i).Value
    modele = ActiveSheet.Range("A" & i).Value '8MODELNENM ou (8MODELZ62M)

    '-------- Barre de recherche --------
    session.findById("wnd[0]/tbar[0]/okcd").Text = "mm01"
    session.findById("wnd[0]").sendVKey 0

    '-------- Cr�er article (Ecran initial) --------
    session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").Text = article 'Article
    session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").Key = "M" 'Branche
    session.findById("wnd[0]/usr/cmbRMMG1-MTART").Key = "CMS" 'Type d'article (CMS - CMS)
    session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").Text = modele 'Mod�le

    'V�rifier l'article pour le site � Nantes ou � Saint Nazaire
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
    session.findById("wnd[1]/tbar[0]/btn[5]").press 'S�lection des vues
    
    'Effacer la selection
    session.findById("wnd[1]/tbar[0]/btn[19]").press
    
    'S�lection des vues
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).Selected = True 'Donn�es de base
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).Selected = True 'Achats
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).Selected = True 'Texte de commande
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).Selected = True 'MRP 1
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).Selected = True 'MRP 2
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).Selected = True 'Donn�es g�n. div./stockage
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).Selected = True 'Gestion emplacements magasin
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(15).Selected = True 'Comptabilit�

    session.findById("wnd[1]/tbar[0]/btn[0]").press
    
    If session.ActiveWindow.Name = "wnd[2]" Then 'Articles cr�� completament
        
        If (session.findById("wnd[2]/usr/txtMESSTXT1").Text = "Article d�j� g�r� pour cette op�ration") Then
            cpt = cpt + 1
            articlesDejaCrees = articlesDejaCrees & article & " "
        End If
        
        session.findById("wnd[2]").Close
        session.findById("wnd[1]").Close
        session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retour
        
    Else 'Articles non finis
    
        compteur = compteur + 1
        articlesIncomplets = articlesIncomplets & article & " "
        
        session.findById("wnd[0]/tbar[0]/btn[15]").press 'Terminer
        session.findById("wnd[1]/usr/btnSPOP-OPTION2").press 'Ne pas sauvegarder des donn�es
    
    End If

Next i

MsgBox ("La v�rification des articles est finie ! Vous avez v�rifi� " & (fin - 3) & " articles." _
& Chr(13) & cpt & " sont d�j� cr��s" & Chr(13) & compteur & " articles sont incomplets ")

If MsgBox("Voulez-vous savoir quels sont les articles � v�rifier ?", vbYesNo + vbQuestion, "Articles incomplets") = vbYes Then
    MsgBox articlesIncomplets
End If

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If

End Sub
