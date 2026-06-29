Sub Remplir_Etat()

    Dim wsP As Worksheet
    Dim wsE As Worksheet

    Dim i As Long
    Dim derLigP As Long
    Dim derLigE As Long

    Dim Jour As String
    Dim Ville As String
    Dim Heure As String

    Set wsP = Worksheets("PLANNING")
    Set wsE = Worksheets("ETAT")

    derLigP = wsP.Cells(wsP.Rows.Count, "A").End(xlUp).Row
    derLigE = wsE.Cells(wsE.Rows.Count, "A").End(xlUp).Row

    'Effacer les anciens résultats
    wsE.Range("D6:I34").ClearContents
    wsE.Range("N6:X34").ClearContents

    Dim colEntree As Long
    Dim colSortie As Long

    For i = 6 To derLigE

        Jour = UCase(wsE.Cells(i, 1).Value)
        Ville = UCase(wsE.Cells(i, 2).Value)

        Select Case Jour

            Case "LUNDI"
                colEntree = 12
                colSortie = 13

            Case "MARDI"
                colEntree = 14
                colSortie = 15

            Case "MERCREDI"
                colEntree = 16
                colSortie = 17

            Case "JEUDI"
                colEntree = 18
                colSortie = 19

            Case "VENDREDI"
                colEntree = 20
                colSortie = 21

            Case "SAMEDI"
                colEntree = 22
                colSortie = 23

            Case "DIMANCHE"
                colEntree = 24
                colSortie = 25

        End Select

        Dim c As Long

        'Entrées
        For c = 4 To 8

            Heure = Format(wsE.Cells(5, c).Value, "00") & ":00"

            wsE.Cells(i, c).Value = WorksheetFunction.CountIfs( _
                    wsP.Range("I:I"), Ville, _
                    wsP.Columns(colEntree), Heure)

        Next c

        'Sorties
        For c = 14 To 23

            Heure = Format(wsE.Cells(5, c).Value, "00") & ":00"

            wsE.Cells(i, c).Value = WorksheetFunction.CountIfs( _
                    wsP.Range("I:I"), Ville, _
                    wsP.Columns(colSortie), Heure)

        Next c

    Next i

    MsgBox "Etat mis à jour."

End Sub
