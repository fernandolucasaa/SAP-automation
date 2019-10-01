Attribute VB_Name = "recupererEmplacements"
Option Explicit

'Recuperer les emplacements de tous les articles dans la liste

Sub recupererEmplacement_SAP()

logonSAP 'Se connecter au SAP

'_________________________________________________________________________________________________'
                    'Récupérer les emplacements

Dim article As String, fichier As String, fin As String
Dim i As Integer, j As Integer, compteur As Integer
Dim ws As Worksheet

fichier = ThisWorkbook.Name
Set ws = Workbooks(fichier).Worksheets("Feuil1")

fin = ws.Cells(Rows.Count, 1).End(xlUp).Row

compteur = 0

'Boucle pour chercher des données
For i = 2 To fin
    
    session.startTransaction ("MM03")
    
    '-------- Afficher Article (Ecran initial) -------
    article = ws.Range("A" & i).Value
    userArea.findById("ctxtRMMG1-MATNR").Text = article
    
    Dim division As String, magasin As String, numeroMagasin As String, typeMagasin As String

    division = ws.Range("C" & i).Value 'NTF ou (NZF)
    magasin = ws.Range("D" & i).Value 'NENM ou (Z62M)
    numeroMagasin = ws.Range("E" & i).Value 'N18 ou (Z18)
    typeMagasin = ws.Range("F" & i).Value 'NEN ou (Z62)
    
    wnd0.findById("tbar[1]/btn[6]").press 'ouvrir le "Niveaux de organization"
    
    'Configurer le niveau de organization
    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = "" 'Division
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = "" 'Magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").Text = "" 'Numero magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").Text = "" 'Type magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").Text = division
    session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").Text = magasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").Text = numeroMagasin
    session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").Text = typeMagasin
    session.findById("wnd[1]/tbar[0]/btn[5]").press 'Sélection des vues

    'Sélection des vues
    session.findById("wnd[1]/tbar[0]/btn[19]").press 'Demarquer tout
    
    Do While session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0," & j & "]").Text <> "Gestion emplacements magasin"
        j = j + 1
        'Debug.Print session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0," & j & "]").Text
    Loop
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(j).Selected = True 'Gestion emplacements magasin
    j = 0
    
    session.findById("wnd[1]/tbar[0]/btn[0]").press 'Suite
    
    '-------- Afficher Article (Données gén. div./stockage, CMS - CMS) --------
    Dim emplacement As String
    
    emplacement = session.findById("wnd[0]/usr/subSUB5:SAPLMGD1:2734/ctxtMLGT-LGPLA").Text
    ws.Range("B" & i).Value = emplacement
    
    wnd0.sendVKey 0 'Enter
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press 'Quitter l'affichade de l'article
    session.findById("wnd[0]/tbar[0]/btn[3]").press 'Retourner à le premier écran
    
    compteur = compteur + 1

Next i
    
'Sauvegarder
Workbooks(fichier).Save

MsgBox ("Vous avez récupéré l'emplacement des " & compteur & " articles !")

'Fermeture de la connexion
If MsgBox("Voulez-vous fermer votre session SAP ?", vbYesNo, "Fermeture de la session SAP") = vbYes Then
    fermetureSAP
End If
    
End Sub

