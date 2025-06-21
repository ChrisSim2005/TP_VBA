'module 1
Sub Creation_Tableau_Facture()
    Dim feuille_active As Worksheet
    
    'selection la feuile active (Facture)
    Set feuille_active = Sheets("FACTURE")
    
    ' Effacer le contenu existant
    feuille_active.Cells.Clear
    
    ' Création des en-têtes
    With feuille_active
        .Range("B3").Value = "Produit"
        .Range("C3").Value = "Quantité"
        .Range("D3").Value = "Prix unitaire"
        .Range("E3").Value = "Montant"
        
        ' Mise en forme des en-têtes produit , quantité , prix unitaire , montant
        With .Range("B3:E3")
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
        ' Ajustement des colonnes
        Columns("B").ColumnWidth = 20
        Columns("C:E").ColumnWidth = 15
        
        ' Application des couleurs de fond
        Range("C4:C10").Interior.Color = RGB(255, 255, 0)
        Range("D4:D10").Interior.Color = RGB(0, 255, 0)
        Range("E4:E10").Interior.Color = RGB(255, 0, 0)
        
        ' Ajout des bordures
        With .Range("B3:E10")
            .Borders.LineStyle = xlContinuous
        End With
    End With
    
    MsgBox "Facture créée avec succès sur la feuille FACTURE "
End Sub


Sub GenererFacture()
    Dim feuille_active As Worksheet
    Dim calculfacture As Worksheet
    Dim ligne As Integer
    Dim produit As String
    Dim quantite As Double
    Dim prix As Double
    Dim i As Integer
    
    'inisalization de ligne et i
    ligne = 0
    i = 0
    'selection la feuile active (Facture)
    Set feuille_active = Sheets("FACTURE")
    'selection la feuile calculactive
    Set calculfacture = Sheets("CalculFacture")
    
    'supprimer si il y a un élèment présent dans les cellules
    With feuille_active
       .Range("B4:B10") = ""
       .Range("C4:C10") = ""
       .Range("D4:D10") = ""
       .Range("E4:E10") = ""
    End With
    
    For ligne = 4 To 10
        ' Saisie du produit
        Do
            produit = InputBox("Entrez le nom du produit " & ligne - 3, "Produit")
            
            'le boutton annuler appuyer
            If StrPtr(produit) = 0 Then
             Exit Sub
            End If
            'Vérifie que c'est des lettres
            If IsNumeric(produit) Then
                MsgBox "erreur Seules les lettres alphabétiques sont autorisées"
            Else
                Exit Do
            End If
            
        Loop
        
        ' Saisie de la quantité
        Do
            quantite = InputBox("Quantité pour " & produit, "Quantité")
            
            'le boutton annuler appuyer
            If StrPtr(quantite) = 0 Then
             Exit Sub
            End If
            
            ' Vérification que c'est un nombre valide
            If IsNumeric(quantite) And quantite > 0 Then
                Exit Do
            Else
                MsgBox "Erreur : Seuls les nombres positifs sont autorisés"
            End If
        Loop
        
        
        ' Saisie du prix
        Do
            prix = InputBox("Prix unitaire pour " & produit, "Prix")
            
            'le boutton annuler appuyer
            If StrPtr(prix) = 0 Then
             Exit Sub
            End If
            
            ' Vérification que c'est un nombre valide
            If IsNumeric(prix) And prix > 0 Then
                Exit Do
            Else
                MsgBox "Erreur : Seuls les nombres positifs sont autorisés"
            End If
        Loop
        
        
        ' Enregistrement dans la feuille facture
        With feuille_active
                .Cells(ligne, 2).Value = produit
                .Cells(ligne, 3).Value = quantite
                .Cells(ligne, 4).Value = prix
                .Cells(ligne, 5).Value = prix * quantite
        End With
        
        ' Format de prix et montant
        With feuille_active
                .Cells(ligne, 3).NumberFormat = "0"
                .Cells(ligne, 4).NumberFormat = "0 €"
                .Cells(ligne, 5).NumberFormat = "0 €"
                
        End With
        i = ligne + 5
        ' Enregistrement dans la feuille calcul facture
        With calculfacture
                .Cells(i, 2).Value = produit
                .Cells(i, 3).Value = quantite
                .Cells(i, 4).Value = prix
                .Cells(i, 3).NumberFormat = "0"
                .Cells(i, 4).NumberFormat = "0 €"
        End With
    
    Next ligne
    
    MsgBox "La saisie est terminée !" 'affiche un message saisie terminée
End Sub

Sub sommeM()
    Dim feuille_active As Worksheet
    Dim plage_montants As Range
    
    ' Utiliser la feuille active
    Set feuille_active = Sheets("FACTURE")
    
    ' Définir la plage des montants (E4:E10)
    Set plage_montants = feuille_active.Range("E4:E10")
    
    ' Ajouter le TOTAL en B16
    With feuille_active.Range("B16")
        .Value = "TOTAL"
        .Font.Bold = True
    End With
    
    ' Calculer le total en E16
    With feuille_active.Range("E16")
        .Formula = "=SUM(" & plage_montants.Address & ")"
        .NumberFormat = "0 €"
        .Font.Bold = True
    End With
    
   
End Sub


'module 2



Sub initialiser()

Set feuille_active = Sheets("CalculFacture")
 With feuille_active
       .Range("B9:B15") = ""
       .Range("C9:C15") = ""
       .Range("D9:D15") = ""
       .Range("E9:E15") = ""
       .Range("F19:F20") = ""
End With
End Sub

Sub calculMontant()
Set feuille_active = Sheets("CalculFacture")
Dim i As Integer
    For i = 9 To 15
        quantite = feuille_active.Cells(i, 3).Value
        prix = feuille_active.Cells(i, 4).Value
        
        If Cells(i, 3).Value <> "" And Cells(i, 4).Value <> "" Then
            With feuille_active
                .Cells(i, 5).Value = prix * quantite
                .Cells(i, 5).NumberFormat = "0 €"
            End With
        Else
            feuille_active.Cells(i, 5) = ""
        End If
    Next i
    
End Sub

Sub CalculerTotalHT()

    Set feuille_active = Sheets("CalculFacture")
    
    Dim i As Integer
    Dim sommeHT As Double
    Dim valeurCellule As Double
    sommeHT = 0

    For i = 9 To 15
        If feuille_active.Cells(i, 5).Value <> "" Then
            valeurCellule = feuille_active.Cells(i, 5).Value
            sommeHT = sommeHT + valeurCellule
        End If
    Next i

    With feuille_active
        .Range("E19").Value = "Total HT :"
        .Cells(19, 6).Value = sommeHT
        .Cells(19, 6).NumberFormat = "0 €"
    End With

End Sub

Sub calculTVA()

Set feuille_active = Sheets("CalculFacture")
Dim TVA As Double

TVA = feuille_active.Range("F19").Value

TVA = TVA * 0.2

With feuille_active
            .Range("E20").Value = "TVA(20%):"
            .Cells(20, 6).Value = TVA
            .Cells(20, 6).NumberFormat = "0 €"
      End With

End Sub

Sub Total_TTC()

Set feuille_active = Sheets("CalculFacture")
Dim TTC As Double
Dim HT As Double
Dim TVA As Double

HT = feuille_active.Range("F19").Value
TVA = feuille_active.Range("F20").Value

TTC = HT + TVA

MsgBox "Total TTC: " & TTC & " €"


End Sub



