Sub Tableau_de_facture()
    Dim feuille As Worksheet
    Dim calculfacture As Worksheet
    Set feuille = Sheets("Facture")
    Set calculfacture = Sheets("Calcul de la Facture")
   
   'nettoyez les feuilles avant l'opération
    
    feuille.Cells.Clear
    calculfacture.Cells.Clear
   'pour créer les colonnes
   
   'sur la la feuille Facture
    With feuille
        .Range("B3").Value = "Produit"
        .Range("C3").Value = "Quantité"
        .Range("D3").Value = "Prix unitaire"
        .Range("E3").Value = "Montant"
        
        With .Range("B3:E3")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        .Columns("B").ColumnWidth = 20
        .Columns("C:E").ColumnWidth = 15
        'pour insérer les couleurs
        
        .Range("C4:C10").Interior.Color = RGB(255, 255, 0)
        .Range("D4:D10").Interior.Color = RGB(0, 255, 0)
        .Range("E4:E10").Interior.Color = RGB(255, 0, 0)
        
        With .Range("B3:E10").Borders
            .LineStyle = xlContinuous
        End With
    End With
    
    'sur la feuille CalculFacture
    
    With calculfacture
        .Range("B8").Value = "Produit"
        .Range("C8").Value = "Quantité"
        .Range("D8").Value = "Prix unitaire"
        .Range("E8").Value = "Montant"
        
        With .Range("B8:E8")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        .Columns("B").ColumnWidth = 20
        .Columns("C:E").ColumnWidth = 15
        
        .Range("C9:C15").Interior.Color = RGB(255, 255, 0)
        .Range("D9:D15").Interior.Color = RGB(0, 255, 0)
        .Range("E9:E15").Interior.Color = RGB(255, 0, 0)
        
        With .Range("B8:E15").Borders
            .LineStyle = xlContinuous
        End With
    End With
    
    MsgBox "Le tableau de facture a été créé sur la feuille"
End Sub

Sub remplissage_du_tableau()
    Dim feuille As Worksheet
    Dim calculfacture As Worksheet
    Dim i As Integer
    Dim j As Integer
    Dim produit As String
    Dim quantite As Variant
    Dim prix As Variant
    Set feuille = Sheets("Facture")
    Set calculfacture = Sheets("Calcul de la Facture")
    
    ' Effacer les éléments dans les cellules
    With feuille
        .Range("B4:E10") = ""
    End With
    With calculfacture
        .Range("B9:E15") = ""
    End With

    ' Remplissage des deux tableaux en meme temps
    For i = 4 To 10
        j = i + 5

        ' Saisie du produit
        Do
            produit = InputBox("Entrez le nom du produit:", "Produit")
            If StrPtr(produit) = 0 Then Exit Sub
            If Not produit Like "*[!A-Za-z]*" And Len(Trim(produit)) >= 2 Then Exit Do
            MsgBox "Veuillez entrer uniquement des lettres ."
        Loop

        ' Saisie de la quantité
        Do
            quantite = InputBox("Entrez la quantité pour le produit: " & produit & " :", "Quantité")
            If StrPtr(quantite) = 0 Then Exit Sub
            If IsNumeric(quantite) And quantite > 0 Then Exit Do
            MsgBox "Veuillez entrer une quantité numérique positive."
        Loop

        ' Saisie du prix
        Do
            prix = InputBox("Entrez le prix unitaire pour  le produit :" & produit & " :", "Prix")
            If StrPtr(prix) = 0 Then Exit Sub
            If IsNumeric(prix) And prix > 0 Then Exit Do
            MsgBox "Veuillez entrer un prix numérique positif."
        Loop

        ' Enregistrement dans la feuille "Facture"
        With feuille
            .Cells(i, 2).Value = produit
            .Cells(i, 3).Value = quantite
            .Cells(i, 4).Value = prix
            .Cells(i, 5).Value = prix * quantite
            .Cells(i, 3).NumberFormat = "0"
            .Cells(i, 4).NumberFormat = "0 €"
            .Cells(i, 5).NumberFormat = "0 €"
        End With

        ' Enregistrement dans "Calcul de la Facture"
        With calculfacture
            .Cells(j, 2).Value = produit
            .Cells(j, 3).Value = quantite
            .Cells(j, 4).Value = prix
            .Cells(j, 3).NumberFormat = "0"
            .Cells(j, 4).NumberFormat = "0 €"
        End With
    Next i

    MsgBox "La saisie est terminée  !"
End Sub


Sub sommeM()
Dim feuille As Worksheet
    Dim plage_montants As Range
    
    ' Utiliser la feuille active
    Set feuille = Sheets("Facture")
    
    ' Définir la plage des montants (E4:E10)
    Set plage_montants = feuille.Range("E4:E10")
    
    ' Ajouter le TOTAL en B16
    With feuille.Range("B16")
        .Value = "TOTAL"
        .Font.Bold = True
    End With
    
    ' Calculer le total en E16
    With feuille.Range("E16")
        .Formula = "=SUM(" & plage_montants.Address & ")"
        .NumberFormat = "0.00 €"
        .Font.Bold = True
    End With

End Sub


