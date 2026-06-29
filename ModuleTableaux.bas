Attribute VB_Name = "ModuleTableaux"
' ============================================================
' MODULE : ModuleTableaux
' Génération de la feuille ETAT :
'   Tableau 1 — Entrées  (lignes=Jour+Ville, colonnes=heures)
'   Tableau 2 — Sorties  (lignes=Jour+Ville, colonnes=heures)
' ============================================================
Option Explicit

' Espacement entre les deux tableaux (nombre de lignes vides)
Private Const ESPACEMENT_TABLEAUX As Integer = 3

' ============================================================
' POINT D'ENTRÉE : construire la feuille ETAT entièrement
' ============================================================
Public Sub GenererFeuille()
    DebutCalcul

    On Error GoTo ErrHandler

    ' 1. Calculer les statistiques
    ModuleStatistiques.CalculerStatistiques

    ' 2. Préparer la feuille ETAT
    Dim ws As Worksheet
    Set ws = ObtenirOuCreerFeuille(NOM_FEUILLE_ETAT)
    ws.Cells.Clear
    ws.Tab.Color = RGB(68, 114, 196)

    ' 3. Récupérer les villes
    Dim villes() As String: villes = ModulePlanning.GetVilles()
    Dim nbVilles As Integer: nbVilles = ModulePlanning.GetNbVilles()
    If nbVilles = 0 Then
        ws.Cells(1, 1).Value = "Aucune donnée trouvée dans PLANNING."
        GoTo Cleanup
    End If

    ' 4. Construire les tableaux
    Dim ligneDebut As Long: ligneDebut = 1

    ' ── Tableau Entrées ──
    Dim ligneFinEntrees As Long
    ligneFinEntrees = GenererTableau(ws, ligneDebut, villes, nbVilles, True)

    ' ── Tableau Sorties ──
    Dim ligneDebutSorties As Long
    ligneDebutSorties = ligneFinEntrees + ESPACEMENT_TABLEAUX
    GenererTableau ws, ligneDebutSorties, villes, nbVilles, False

    ' 5. Figer les volets sur la 2ème colonne + ligne après en-tête
    ws.Activate
    ws.Cells(ligneDebut + 2, 2).Select
    ActiveWindow.FreezePanes = True

    ws.Cells(1, 1).Select

    MsgBox "Statistiques actualisées avec succès !", vbInformation, "ETAT"
    GoTo Cleanup

ErrHandler:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical
Cleanup:
    FinCalcul
End Sub

' ============================================================
' GÉNÉRER UN TABLEAU (entrées OU sorties)
' Retourne le numéro de la dernière ligne utilisée
' ============================================================
Private Function GenererTableau(ws As Worksheet, ligneDebut As Long, _
                                 villes() As String, nbVilles As Integer, _
                                 estEntrees As Boolean) As Long
    Dim hMin As Integer, hMax As Integer
    Dim titre As String, couleurTitre As Long

    If estEntrees Then
        hMin = HEURE_MIN_ENTREE: hMax = HEURE_MAX_ENTREE
        titre = "ENTRÉES — Nombre d'agents par heure de prise de poste"
        couleurTitre = RGB(31, 73, 125)
    Else
        hMin = HEURE_MIN_SORTIE: hMax = HEURE_MAX_SORTIE
        titre = "SORTIES — Nombre d'agents par heure de fin de service"
        couleurTitre = RGB(120, 40, 10)
    End If

    Dim nbHeures As Integer: nbHeures = hMax - hMin + 1
    Dim nbJours As Integer:  nbJours  = NB_JOURS

    ' ── Ligne 1 : titre du tableau ──
    Dim rTitre As Range
    Set rTitre = ws.Range(ws.Cells(ligneDebut, 1), ws.Cells(ligneDebut, 1 + nbHeures))
    rTitre.Merge
    rTitre.Value = titre
    With rTitre
        .Font.Bold = True
        .Font.Size = 13
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = couleurTitre
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 28
    End With

    ' ── Ligne 2 : en-têtes colonnes ──
    Dim ligneEnTete As Long: ligneEnTete = ligneDebut + 1
    ws.Cells(ligneEnTete, 1).Value = "Jour / Ville"
    StyleEnTete ws.Cells(ligneEnTete, 1)
    ws.Columns(1).ColumnWidth = 26

    Dim h As Integer
    For h = hMin To hMax
        Dim colH As Integer: colH = 2 + (h - hMin)
        ws.Cells(ligneEnTete, colH).Value = FormatHeure(h)
        StyleEnTete ws.Cells(ligneEnTete, colH)
        ws.Columns(colH).ColumnWidth = 7
    Next h

    ' Colonne Totaux
    Dim colTot As Integer: colTot = 2 + nbHeures
    ws.Cells(ligneEnTete, colTot).Value = "TOTAL"
    With ws.Cells(ligneEnTete, colTot)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(80, 80, 80)
        .HorizontalAlignment = xlCenter
    End With
    ws.Columns(colTot).ColumnWidth = 9

    ws.Rows(ligneEnTete).RowHeight = 22

    ' ── Lignes de données : Jour + Ville ──
    Dim ligneData As Long: ligneData = ligneEnTete + 1
    Dim j As Integer, v As Integer

    For j = 1 To nbJours
        Dim nomJ As String: nomJ = NomJour(j)
        Dim bgJour As Long: bgJour = CouleurJour(j)

        For v = 1 To nbVilles
            Dim ville As String: ville = villes(v)

            ' Libellé de ligne
            Dim libelle As String: libelle = nomJ & "  —  " & ville
            ws.Cells(ligneData, 1).Value = libelle
            With ws.Cells(ligneData, 1)
                .Font.Bold = True
                .Font.Color = RGB(255, 255, 255)
                .Interior.Color = bgJour
                .HorizontalAlignment = xlLeft
                .IndentLevel = 1
            End With

            ' Valeurs par heure + calcul du total ligne
            Dim totalLigne As Integer: totalLigne = 0
            For h = hMin To hMax
                Dim colC As Integer: colC = 2 + (h - hMin)
                Dim val As Integer

                If estEntrees Then
                    val = ModuleStatistiques.GetValeurEntree(nomJ, ville, h)
                Else
                    val = ModuleStatistiques.GetValeurSortie(nomJ, ville, h)
                End If

                ws.Cells(ligneData, colC).Value = IIf(val = 0, "", val)
                ws.Cells(ligneData, colC).HorizontalAlignment = xlCenter

                If val > 0 Then
                    ' Couleur dégradée selon l'intensité
                    ws.Cells(ligneData, colC).Interior.Color = CouleurIntensity(val, estEntrees)
                    ws.Cells(ligneData, colC).Font.Bold = True
                    ws.Cells(ligneData, colC).Font.Color = RGB(0, 0, 0)
                Else
                    ws.Cells(ligneData, colC).Interior.Color = COULEUR_ZERO
                    ws.Cells(ligneData, colC).Font.Color = RGB(180, 180, 180)
                End If

                totalLigne = totalLigne + val
            Next h

            ' Cellule total
            ws.Cells(ligneData, colTot).Value = IIf(totalLigne = 0, "", totalLigne)
            With ws.Cells(ligneData, colTot)
                .HorizontalAlignment = xlCenter
                .Font.Bold = True
                If totalLigne > 0 Then
                    .Interior.Color = RGB(80, 80, 80)
                    .Font.Color = RGB(255, 255, 255)
                Else
                    .Interior.Color = COULEUR_ZERO
                    .Font.Color = RGB(180, 180, 180)
                End If
            End With

            ws.Rows(ligneData).RowHeight = 18
            ligneData = ligneData + 1
        Next v

        ' Ligne séparatrice entre les jours
        Dim rSep As Range
        Set rSep = ws.Range(ws.Cells(ligneData, 1), ws.Cells(ligneData, colTot))
        rSep.Interior.Color = RGB(200, 200, 200)
        rSep.RowHeight = 4
        ligneData = ligneData + 1
    Next j

    ' Ligne de totaux colonnes (totaux par heure)
    ws.Cells(ligneData, 1).Value = "TOTAL / Heure"
    With ws.Cells(ligneData, 1)
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(80, 80, 80)
        .HorizontalAlignment = xlCenter
    End With

    Dim grandTotal As Long: grandTotal = 0
    For h = hMin To hMax
        colC = 2 + (h - hMin)
        Dim totalCol As Long: totalCol = 0
        For j = 1 To nbJours
            Dim nomJT As String: nomJT = NomJour(j)
            For v = 1 To nbVilles
                If estEntrees Then
                    totalCol = totalCol + ModuleStatistiques.GetValeurEntree(nomJT, villes(v), h)
                Else
                    totalCol = totalCol + ModuleStatistiques.GetValeurSortie(nomJT, villes(v), h)
                End If
            Next v
        Next j
        ws.Cells(ligneData, colC).Value = IIf(totalCol = 0, "", totalCol)
        With ws.Cells(ligneData, colC)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(80, 80, 80)
            .Font.Color = RGB(255, 255, 255)
        End With
        grandTotal = grandTotal + totalCol
    Next h

    ws.Cells(ligneData, colTot).Value = grandTotal
    With ws.Cells(ligneData, colTot)
        .Font.Bold = True: .Font.Color = RGB(255, 255, 0)
        .Interior.Color = RGB(50, 50, 50)
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(ligneData).RowHeight = 20

    ' Bordures sur la plage complète
    AppliquerBordures ws.Range(ws.Cells(ligneEnTete, 1), ws.Cells(ligneData, colTot))

    GenererTableau = ligneData
End Function

' ============================================================
' COULEUR PAR JOUR (nuances de bleu/vert/orange)
' ============================================================
Private Function CouleurJour(j As Integer) As Long
    Select Case j
        Case 1: CouleurJour = RGB(31, 73, 125)   ' Lundi — bleu foncé
        Case 2: CouleurJour = RGB(0, 112, 192)   ' Mardi — bleu
        Case 3: CouleurJour = RGB(0, 176, 240)   ' Mercredi — bleu clair
        Case 4: CouleurJour = RGB(0, 128, 0)     ' Jeudi — vert foncé
        Case 5: CouleurJour = RGB(112, 173, 71)  ' Vendredi — vert clair
        Case 6: CouleurJour = RGB(197, 90, 17)   ' Samedi — orange foncé
        Case 7: CouleurJour = RGB(255, 192, 0)   ' Dimanche — jaune/or
        Case Else: CouleurJour = RGB(100, 100, 100)
    End Select
End Function

' ============================================================
' COULEUR D'INTENSITÉ selon la valeur
' Entrées : dégradé vert clair → vert foncé
' Sorties  : dégradé rose → rouge
' ============================================================
Private Function CouleurIntensity(val As Integer, estEntrees As Boolean) As Long
    Dim ratio As Double
    ratio = Application.WorksheetFunction.Min(1, val / 50)  ' Max référence 50 agents

    If estEntrees Then
        ' Vert clair (197,224,180) → vert foncé (0,97,0)
        Dim r1 As Integer: r1 = 197 - Int(197 * ratio)
        Dim g1 As Integer: g1 = 224 - Int(127 * ratio)
        Dim b1 As Integer: b1 = 180 - Int(180 * ratio)
        CouleurIntensity = RGB(r1, g1, b1)
    Else
        ' Rose clair (255,230,200) → rouge (192,0,0)
        Dim r2 As Integer: r2 = 255 - Int(63 * ratio)
        Dim g2 As Integer: g2 = 230 - Int(230 * ratio)
        Dim b2 As Integer: b2 = 200 - Int(200 * ratio)
        CouleurIntensity = RGB(r2, g2, b2)
    End If
End Function
