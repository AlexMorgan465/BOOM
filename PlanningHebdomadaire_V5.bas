Sub Remplir_Etat()

    Dim wsP As Worksheet
    Dim wsE As Worksheet
    Dim i As Long, c As Long
    Dim derLigE As Long

    Dim Jour As String
    Dim Ville As String
    Dim Heure As Variant

    Dim colEntree As Long
    Dim colSortie As Long

    Set wsP = Worksheets("PLANNING")
    Set wsE = Worksheets("ETAT")

    derLigE = wsE.Cells(wsE.Rows.Count, "A").End(xlUp).Row

    'Effacer les anciens résultats
    wsE.Range("C6:I" & derLigE).ClearContents
    wsE.Range("N6:W" & derLigE).ClearContents

    For i = 6 To derLigE

        Jour = UCase(Trim(wsE.Cells(i, 1).Value))
        Ville = UCase(Trim(wsE.Cells(i, 2).Value))

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

        '=====================
        ' ENTRÉES
        '=====================
        For c = 3 To 9

            Heure = wsE.Cells(5, c).Value

            If Ville = "SALE SALA AL JADIDA" Then

                wsE.Cells(i, c).Value = _
                    WorksheetFunction.CountIfs(wsP.Range("I:I"), "SALE", wsP.Columns(colEntree), Heure) + _
                    WorksheetFunction.CountIfs(wsP.Range("I:I"), "SALA AL JADIDA", wsP.Columns(colEntree), Heure)

            Else

                wsE.Cells(i, c).Value = WorksheetFunction.CountIfs( _
                    wsP.Range("I:I"), Ville, _
                    wsP.Columns(colEntree), Heure)

            End If

        Next c

        '=====================
        ' SORTIES
        '=====================
        For c = 14 To 23

            Heure = wsE.Cells(5, c).Value

            If Ville = "SALE SALA AL JADIDA" Then

                wsE.Cells(i, c).Value = _
                    WorksheetFunction.CountIfs(wsP.Range("I:I"), "SALE", wsP.Columns(colSortie), Heure) + _
                    WorksheetFunction.CountIfs(wsP.Range("I:I"), "SALA AL JADIDA", wsP.Columns(colSortie), Heure)

            Else

                wsE.Cells(i, c).Value = WorksheetFunction.CountIfs( _
                    wsP.Range("I:I"), Ville, _
                    wsP.Columns(colSortie), Heure)

            End If

        Next c

    Next i

    MsgBox "Mise à jour de la feuille ETAT terminée.", vbInformation

End Sub
