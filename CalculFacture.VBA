Sub initialiser()
    Dim feuille1 As Worksheet
    Dim feuille2 As Worksheet

    Set feuille1 = Sheets("Calcul de la Facture")
   'pour initialiser le contenu de ces plages de cellules
   
    With feuille1
        .Range("C9:C15") = ""
        .Range("D9:D15") = ""
        .Range("E9:E15") = ""
    End With
End Sub
        

Sub calculer_total_HT()
    Dim feuille As Worksheet
    Dim totalHT As Double
    Dim i As Integer

    Set feuille = Sheets("Calcul de la Facture")
    totalHT = 0

    For i = 9 To 15
        If IsNumeric(feuille.Cells(i, 5).Value) Then
            totalHT = totalHT + feuille.Cells(i, 5).Value
        End If
    Next i

    feuille.Range("F19").Value = "Total HT :"
    feuille.Range("G19").Value = totalHT
    feuille.Range("G19").NumberFormat = "0 €"
    feuille.Range("F19").Font.Bold = True
    feuille.Range("G19").Font.Bold = True
End Sub

Sub calculer_TVA()
    Dim feuille As Worksheet
    Dim totalHT As Double
    Dim montantTVA As Double
    Dim i As Integer

    Set feuille = Sheets("Calcul de la Facture")

    totalHT = Range("G19").Value
    ' la formule de la TVA
    montantTVA = totalHT * 0.2

    feuille.Range("F20").Value = "TVA (20%) :"
    feuille.Range("G20").Value = montantTVA
    feuille.Range("G20").NumberFormat = "0 €"
    feuille.Range("F20").Font.Bold = True
    feuille.Range("G20").Font.Bold = True
End Sub

Sub total_TTC()
    Dim feuille As Worksheet
    Dim totalHT As Double
    Dim totalTTC As Double
    Dim i As Integer

    Set feuille = Sheets("Calcul de la Facture")
    totalHT = Range("G19").Value
        TVA = Range("G20").Value
    'la formule du TTC
    totalTTC = totalHT + TVA

    MsgBox "Le montant total TTC est de : " & Format(totalTTC, "0 €"), vbInformation, "Total TTC"
End Sub

Sub caluculer_les_montants()
    Dim feuille As Worksheet
    Dim i As Integer
    Dim quantite As Variant
    Dim prix As Variant

    Set feuille = Sheets("Calcul de la Facture")
  'la formiule du montant
  
    For i = 9 To 15
        quantite = feuille.Cells(i, 3).Value
        prix = feuille.Cells(i, 4).Value
       'pour remplir les colonnes avec les montants
       
        If IsNumeric(quantite) And IsNumeric(prix) Then
            feuille.Cells(i, 5).Value = quantite * prix
            feuille.Cells(i, 3).NumberFormat = "0"
            feuille.Cells(i, 4).NumberFormat = "0 €"
            feuille.Cells(i, 5).NumberFormat = "0 €"
        Else
            With feuille
                .Range("E9:E15") = ""
             End With
        End If
    Next i
End Sub





