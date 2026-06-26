' ============================================================
' GÉNÉRATEUR AUTOMATIQUE DE PLANNING HEBDOMADAIRE - V5
' NOUVEAUTÉS V5 :
'   - Champ Congé (Oui/Non), Congé D (date début), Congé F (date fin)
'     dans feuille Utilisateurs (colonnes 5,6,7)
'   - Congé = cellule "CONGE" (orange) dans les plannings projet
'   - Congé compte comme OFF dans le compteur quotidien (Google Leads, TLV)
'   - Google Leads : objectif OFF réduit à ~5/jour (+ congés = OFF effectif)
'     pondération Dim/Sam renforcée pour les OFF planifiés
'   - Feuille CONSOLIDATION générée automatiquement :
'     Nom | Date | Entrée | Sortie | Pause D | Pause F |
'     No Semaine | Activité | Congé D | Congé F | Ville | Zone
' ============================================================

Option Explicit

Type Collaborateur
    NomComplet  As String
    projet      As String
    ville       As String
    zone        As String
    IndexRotation As Integer
    EnConge     As Boolean
    CongeDebut  As Date
    CongeFin    As Date
End Type

Dim JOURS(1 To 7) As String

' ============================================================
' POINT D'ENTRÉE PRINCIPAL
' ============================================================
Sub GenererPlanning()
    JOURS(1) = "Lundi"
    JOURS(2) = "Mardi"
    JOURS(3) = "Mercredi"
    JOURS(4) = "Jeudi"
    JOURS(5) = "Vendredi"
    JOURS(6) = "Samedi"
    JOURS(7) = "Dimanche"

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    If Not VerifierFeuillesExistantes() Then
        MsgBox "Erreur : certaines feuilles requises sont manquantes.", vbCritical
        GoTo Cleanup
    End If

    InitialiserFeuilleRotation
    InitialiserFeuilleConsolidation
    EffacerAnciensPlannings

    Dim collaborateurs() As Collaborateur
    Dim nbCollab As Integer
    nbCollab = LireCollaborateurs(collaborateurs)

    If nbCollab = 0 Then
        MsgBox "Aucun collaborateur trouve dans la feuille Utilisateurs."
        GoTo Cleanup
    End If

    GenererPlanningAFEDIM       collaborateurs, nbCollab
    GenererPlanningACCESSIBILITE collaborateurs, nbCollab
    GenererPlanningCMLEASING    collaborateurs, nbCollab
    GenererPlanningGLF          collaborateurs, nbCollab
    GenererPlanningEBRA         collaborateurs, nbCollab
    GenererPlanningGOOGLELEADS  collaborateurs, nbCollab
    GenererPlanningTLV          collaborateurs, nbCollab
    GenererPlanningFACTO        collaborateurs, nbCollab
    GenererPlanningDAC          collaborateurs, nbCollab
    MettreAJourRotation         collaborateurs, nbCollab

    MsgBox "Planning genere avec succes !" & Chr(10) & _
           "Semaine " & Application.WorksheetFunction.WeekNum(Date, 2) & " - " & Year(Date), _
           vbInformation, "Generation Planning"

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "Erreur"
    Resume Cleanup
End Sub

' ============================================================
' VÉRIFICATION / UTILITAIRES
' ============================================================
Function VerifierFeuillesExistantes() As Boolean
    Dim req() As String
    req = Split("Utilisateurs,AFEDIM,ACCESSIBILITE,CM Leasing,GLF,EBRA,GOOGLE LEADS,TLV,FACTO,DAC,CONSOLIDATION", ",")
    Dim i As Integer
    For i = 0 To UBound(req)
        If Not FeuilleExiste(req(i)) Then
            MsgBox "Feuille manquante : [" & req(i) & "]", vbCritical
            VerifierFeuillesExistantes = False
            Exit Function
        End If
    Next i
    VerifierFeuillesExistantes = True
End Function

Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not (ws Is Nothing)
End Function

Function AjouterMinutes(heure As String, minutes As Integer) As String
    If heure = "" Then AjouterMinutes = "": Exit Function
    Dim p() As String
    p = Split(heure, ":")
    Dim t As Integer
    t = CInt(p(0)) * 60 + CInt(p(1)) + minutes
    If t < 0 Then t = 0
    AjouterMinutes = Format(t \ 60, "00") & ":" & Format(t Mod 60, "00")
End Function

' Retourne le numéro du lundi de la semaine courante
Function LundiSemaine() As Date
    Dim d As Date
    d = Date
    Dim wd As Integer
    wd = Weekday(d, vbMonday)   ' 1=Lun .. 7=Dim
    LundiSemaine = d - (wd - 1)
End Function

' Date du jour j (1=Lun..7=Dim) pour la semaine courante
Function DateDuJour(j As Integer) As Date
    DateDuJour = LundiSemaine() + (j - 1)
End Function

' Vérifie si un collab est en congé un jour donné (date)
Function EstEnConge(c As Collaborateur, d As Date) As Boolean
    If Not c.EnConge Then EstEnConge = False: Exit Function
    EstEnConge = (d >= c.CongeDebut And d <= c.CongeFin)
End Function

' Formater le contenu d'une cellule jour (prend en compte congé)
Function FormatCelluleJour(debut As String, fin As String, pD As String, pF As String) As String
    If debut = "OFF" Or debut = "" Then
        FormatCelluleJour = "OFF"
        Exit Function
    End If
    Dim s As String
    s = debut & " - " & fin
    If pD <> "" Then
        s = s & Chr(10) & "Pause: " & pD & "-" & pF
    End If
    FormatCelluleJour = s
End Function

' ============================================================
' EN-TÊTE HORIZONTALE
' ============================================================
Sub EcrireEnTeteHorizontale(ws As Worksheet, projet As String)
    ws.Cells(1, 1).Value = "PLANNING HEBDOMADAIRE - " & UCase(projet)
    With ws.Cells(1, 1)
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 10)).Merge

    ws.Cells(2, 1).Value = "Semaine " & Application.WorksheetFunction.WeekNum(Date, 2) & _
                            " | Generee le " & Format(Date, "dd/mm/yyyy")
    ws.Cells(2, 1).Font.Italic = True
    ws.Range(ws.Cells(2, 1), ws.Cells(2, 10)).Merge

    ws.Cells(3, 1).Value = "Collaborateur"
    ws.Cells(3, 2).Value = "Ville"
    ws.Cells(3, 3).Value = "Zone"
    Dim j As Integer
    For j = 1 To 7
        ws.Cells(3, 3 + j).Value = JOURS(j)
    Next j

    With ws.Rows(3)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ws.Columns("A").ColumnWidth = 28
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 14
    Dim c As Integer
    For c = 4 To 10
        ws.Columns(c).ColumnWidth = 22
    Next c
    ws.Rows(3).RowHeight = 20
End Sub

' ============================================================
' ÉCRIRE UNE LIGNE HORIZONTALE (avec gestion congé)
' cellules(j) peut être "OFF", "CONGE", ou "HH:MM - HH:MM\nPause:..."
' ============================================================
Sub EcrireLigneHorizontale(ws As Worksheet, ligne As Integer, nom As String, _
                            ville As String, zone As String, cellules() As String)
    ws.Cells(ligne, 1).Value = nom
    ws.Cells(ligne, 2).Value = ville
    ws.Cells(ligne, 3).Value = zone

    Dim j As Integer
    For j = 1 To 7
        Dim cel As Range
        Set cel = ws.Cells(ligne, 3 + j)
        cel.Value = cellules(j)
        cel.HorizontalAlignment = xlCenter
        cel.VerticalAlignment = xlCenter
        cel.WrapText = True

        Select Case cellules(j)
            Case "OFF"
                cel.Interior.Color = RGB(255, 199, 206)     ' rouge clair
                cel.Font.Bold = True
                cel.Font.Color = RGB(192, 0, 0)
            Case "CONGE"
                cel.Interior.Color = RGB(255, 192, 0)       ' orange/jaune
                cel.Font.Bold = True
                cel.Font.Color = RGB(0, 0, 0)
            Case Else
                cel.Font.Color = RGB(0, 0, 0)
                If ligne Mod 2 = 0 Then
                    cel.Interior.Color = RGB(235, 241, 255)
                Else
                    cel.Interior.Color = RGB(255, 255, 255)
                End If
        End Select
    Next j

    If ligne Mod 2 = 0 Then
        ws.Cells(ligne, 1).Interior.Color = RGB(235, 241, 255)
        ws.Cells(ligne, 2).Interior.Color = RGB(235, 241, 255)
        ws.Cells(ligne, 3).Interior.Color = RGB(235, 241, 255)
    End If
    ws.Rows(ligne).RowHeight = 40
End Sub

Sub AppliquerBorduresH(ws As Worksheet, ligneDebut As Integer, ligneFin As Integer)
    If ligneFin < ligneDebut Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(ligneDebut, 1), ws.Cells(ligneFin, 10))
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(189, 189, 189)
    End With
    With rng.Borders(xlEdgeLeft)
        .Weight = xlMedium: .Color = RGB(68, 114, 196)
    End With
    With rng.Borders(xlEdgeRight)
        .Weight = xlMedium: .Color = RGB(68, 114, 196)
    End With
    With rng.Borders(xlEdgeTop)
        .Weight = xlMedium: .Color = RGB(68, 114, 196)
    End With
    With rng.Borders(xlEdgeBottom)
        .Weight = xlMedium: .Color = RGB(68, 114, 196)
    End With
End Sub

' ============================================================
' CONSOLIDATION - Initialisation de la feuille
' Colonnes : Nom | Date | Entrée | Sortie | Pause D | Pause F
'            No Semaine | Activité | Congé D | Congé F | Ville | Zone
' ============================================================
Sub InitialiserFeuilleConsolidation()
    Dim ws As Worksheet
    If Not FeuilleExiste("CONSOLIDATION") Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "CONSOLIDATION"
    Else
        Set ws = ThisWorkbook.Sheets("CONSOLIDATION")
    End If
    ws.Cells.Clear

    ' En-têtes
    Dim headers As Variant
    headers = Array("Nom", "Date", "Entree", "Sortie", "Pause D", "Pause F", _
                    "No Semaine", "Activite", "Conge D", "Conge F", "Ville", "Zone")
    Dim c As Integer
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Largeurs colonnes
    ws.Columns("A").ColumnWidth = 28   ' Nom
    ws.Columns("B").ColumnWidth = 14   ' Date
    ws.Columns("C").ColumnWidth = 10   ' Entrée
    ws.Columns("D").ColumnWidth = 10   ' Sortie
    ws.Columns("E").ColumnWidth = 10   ' Pause D
    ws.Columns("F").ColumnWidth = 10   ' Pause F
    ws.Columns("G").ColumnWidth = 12   ' No Semaine
    ws.Columns("H").ColumnWidth = 16   ' Activité
    ws.Columns("I").ColumnWidth = 14   ' Congé D
    ws.Columns("J").ColumnWidth = 14   ' Congé F
    ws.Columns("K").ColumnWidth = 14   ' Ville
    ws.Columns("L").ColumnWidth = 14   ' Zone

    ' Figer la ligne d'en-tête
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
End Sub

' Ajouter une ligne dans CONSOLIDATION
' activite = projet | entree/sortie/pD/pF = "" si OFF ou CONGE
Sub AjouterLigneConsolidation(collab As Collaborateur, d As Date, _
                               entree As String, sortie As String, _
                               pD As String, pF As String, _
                               activite As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("CONSOLIDATION")
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(d, 2)

    ws.Cells(lr, 1).Value = collab.NomComplet
    ws.Cells(lr, 2).Value = d
    ws.Cells(lr, 2).NumberFormat = "dd/mm/yyyy"
    ws.Cells(lr, 3).Value = entree
    ws.Cells(lr, 4).Value = sortie
    ws.Cells(lr, 5).Value = pD
    ws.Cells(lr, 6).Value = pF
    ws.Cells(lr, 7).Value = sem
    ws.Cells(lr, 8).Value = activite

    ' Congé D / Congé F
    If collab.EnConge Then
        ws.Cells(lr, 9).Value = collab.CongeDebut
        ws.Cells(lr, 9).NumberFormat = "dd/mm/yyyy"
        ws.Cells(lr, 10).Value = collab.CongeFin
        ws.Cells(lr, 10).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 9).Value = ""
        ws.Cells(lr, 10).Value = ""
    End If

    ws.Cells(lr, 11).Value = collab.ville
    ws.Cells(lr, 12).Value = collab.zone

    ' Colorer les lignes OFF/CONGE dans la consolidation
    Select Case activite
        Case "OFF"
            ws.Rows(lr).Interior.Color = RGB(255, 199, 206)
        Case "CONGE"
            ws.Rows(lr).Interior.Color = RGB(255, 230, 153)
        Case Else
            If lr Mod 2 = 0 Then
                ws.Rows(lr).Interior.Color = RGB(235, 241, 255)
            Else
                ws.Rows(lr).Interior.Color = RGB(255, 255, 255)
            End If
    End Select
End Sub

' ============================================================
' ROTATION / GESTION
' ============================================================
Sub InitialiserFeuilleRotation()
    Dim ws As Worksheet
    If Not FeuilleExiste("ROTATION") Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "ROTATION"
    Else
        Set ws = ThisWorkbook.Sheets("ROTATION")
    End If
    If ws.Cells(1, 1).Value = "" Then
        ws.Cells(1, 1).Value = "Collaborateur"
        ws.Cells(1, 2).Value = "Projet"
        ws.Cells(1, 3).Value = "Index Rotation"
        ws.Cells(1, 4).Value = "Derniere MAJ"
        ws.Cells(1, 5).Value = "Semaine"
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
        End With
    End If
End Sub

Sub EffacerAnciensPlannings()
    Dim feuilles() As String
    feuilles = Split("AFEDIM,ACCESSIBILITE,CM Leasing,GLF,EBRA,GOOGLE LEADS,TLV,FACTO,DAC", ",")
    Dim i As Integer
    For i = 0 To UBound(feuilles)
        ThisWorkbook.Sheets(feuilles(i)).Cells.Clear
    Next i
    ' Effacer données CONSOLIDATION (garder en-tête)
    Dim wsC As Worksheet
    Set wsC = ThisWorkbook.Sheets("CONSOLIDATION")
    If wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row > 1 Then
        wsC.Range(wsC.Cells(2, 1), wsC.Cells(wsC.Rows.Count, 12)).Clear
    End If
End Sub

' ============================================================
' LECTURE COLLABORATEURS (colonnes 1-7 dans Utilisateurs)
' Col 1=Nom | 2=Projet | 3=Ville | 4=Zone
' Col 5=Congé(Oui/Non) | 6=Congé D | 7=Congé F
' ============================================================
Function LireCollaborateurs(ByRef collabs() As Collaborateur) As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then LireCollaborateurs = 0: Exit Function
    Dim nb As Integer
    nb = lastRow - 1
    ReDim collabs(1 To nb)
    Dim i As Integer
    For i = 1 To nb
        collabs(i).NomComplet   = Trim(ws.Cells(i + 1, 1).Value)
        collabs(i).projet       = Trim(ws.Cells(i + 1, 2).Value)
        collabs(i).ville        = Trim(ws.Cells(i + 1, 3).Value)
        collabs(i).zone         = Trim(ws.Cells(i + 1, 4).Value)
        collabs(i).IndexRotation = LireIndexRotation(collabs(i).NomComplet, collabs(i).projet)

        ' Lecture congé
        Dim congeVal As String
        congeVal = UCase(Trim(ws.Cells(i + 1, 5).Value))
        collabs(i).EnConge = (congeVal = "OUI" Or congeVal = "O" Or congeVal = "YES")

        If collabs(i).EnConge Then
            Dim rawD As Variant: rawD = ws.Cells(i + 1, 6).Value
            Dim rawF As Variant: rawF = ws.Cells(i + 1, 7).Value
            If IsDate(rawD) Then collabs(i).CongeDebut = CDate(rawD) Else collabs(i).CongeDebut = Date
            If IsDate(rawF) Then collabs(i).CongeFin   = CDate(rawF) Else collabs(i).CongeFin   = Date
        End If
    Next i
    LireCollaborateurs = nb
End Function

Function LireIndexRotation(nom As String, projet As String) As Integer
    If Not FeuilleExiste("ROTATION") Then LireIndexRotation = 0: Exit Function
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ROTATION")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = nom And ws.Cells(i, 2).Value = projet Then
            LireIndexRotation = CInt(ws.Cells(i, 3).Value)
            Exit Function
        End If
    Next i
    LireIndexRotation = 0
End Function

Sub MettreAJourRotation(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ROTATION")
    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(Date, 2)
    Dim i As Integer
    For i = 1 To nb
        Dim lr As Long
        lr = TrouverLigneRotation(ws, collabs(i).NomComplet, collabs(i).projet)
        If lr = 0 Then
            lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
            ws.Cells(lr, 1).Value = collabs(i).NomComplet
            ws.Cells(lr, 2).Value = collabs(i).projet
            ws.Cells(lr, 3).Value = 1
        Else
            ws.Cells(lr, 3).Value = CInt(ws.Cells(lr, 3).Value) + 1
        End If
        ws.Cells(lr, 4).Value = Date
        ws.Cells(lr, 5).Value = sem
    Next i
    ws.Columns("A:E").AutoFit
End Sub

Function TrouverLigneRotation(ws As Worksheet, nom As String, projet As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = nom And ws.Cells(i, 2).Value = projet Then
            TrouverLigneRotation = i: Exit Function
        End If
    Next i
    TrouverLigneRotation = 0
End Function

' ============================================================
' HELPER : construire les cellules d'un collab en tenant
'          compte de son congé sur la semaine courante.
' Passe les tableaux par référence et remplace les jours
' concernés par "CONGE".
' ============================================================
Sub AppliquerConges(ByRef cellules() As String, c As Collaborateur)
    Dim j As Integer
    For j = 1 To 7
        Dim d As Date
        d = DateDuJour(j)
        If EstEnConge(c, d) Then
            cellules(j) = "CONGE"
        End If
    Next j
End Sub

' Helper : écrire une ligne + CONSOLIDATION pour un collab fixe
Sub EcrireLigneAvecConsolidation(ws As Worksheet, ligne As Integer, _
                                  c As Collaborateur, cellules() As String, _
                                  entrees() As String, sorties() As String, _
                                  pDs() As String, pFs() As String)
    EcrireLigneHorizontale ws, ligne, c.NomComplet, c.ville, c.zone, cellules

    Dim j As Integer
    For j = 1 To 7
        Dim d As Date
        d = DateDuJour(j)
        Dim activite As String
        Dim entStr As String, sorStr As String, pdStr As String, pfStr As String
        entStr = "": sorStr = "": pdStr = "": pfStr = ""

        Select Case cellules(j)
            Case "CONGE"
                activite = "CONGE"
            Case "OFF"
                activite = "OFF"
            Case Else
                activite = c.projet
                entStr = entrees(j)
                sorStr = sorties(j)
                pdStr  = pDs(j)
                pfStr  = pFs(j)
        End Select

        AjouterLigneConsolidation c, d, entStr, sorStr, pdStr, pfStr, activite
    Next j
End Sub

' ============================================================
' PROJETS FIXES : AFEDIM / ACCESSIBILITE / CM LEASING
' Lun-Jeu 08:00-18:00 | Ven 08:00-17:00 | Sam-Dim OFF
' Pause fixe 13:00-14:00 | Total = 44h/sem
' ============================================================
Sub GenererPlanningFixe(nomFeuille As String, collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomFeuille)
    EcrireEnTeteHorizontale ws, nomFeuille

    Dim ligne As Integer
    ligne = 4
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = UCase(nomFeuille) Then
            Dim cellules(1 To 7) As String
            Dim entrees(1 To 7) As String
            Dim sorties(1 To 7) As String
            Dim pDs(1 To 7) As String
            Dim pFs(1 To 7) As String
            Dim j As Integer

            For j = 1 To 7
                Select Case j
                    Case 1, 2, 3, 4
                        cellules(j) = FormatCelluleJour("08:00", "18:00", "13:00", "14:00")
                        entrees(j) = "08:00": sorties(j) = "18:00"
                        pDs(j) = "13:00": pFs(j) = "14:00"
                    Case 5
                        cellules(j) = FormatCelluleJour("08:00", "17:00", "13:00", "14:00")
                        entrees(j) = "08:00": sorties(j) = "17:00"
                        pDs(j) = "13:00": pFs(j) = "14:00"
                    Case Else
                        cellules(j) = "OFF"
                        entrees(j) = "": sorties(j) = "": pDs(j) = "": pFs(j) = ""
                End Select
            Next j

            ' Appliquer congés (écrase les jours concernés en "CONGE")
            AppliquerConges cellules, collabs(i)

            EcrireLigneAvecConsolidation ws, ligne, collabs(i), cellules, entrees, sorties, pDs, pFs
            ligne = ligne + 1
        End If
    Next i

    If ligne > 4 Then
        ws.Cells(ligne + 1, 1).Value = "Total hebdomadaire : 44h | Pause fixe 13:00-14:00"
        ws.Cells(ligne + 1, 1).Font.Italic = True
        ws.Cells(ligne + 1, 1).Font.Color = RGB(31, 73, 125)
    End If
    AppliquerBorduresH ws, 4, ligne - 1
End Sub

Sub GenererPlanningAFEDIM(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFixe "AFEDIM", collabs, nb
End Sub
Sub GenererPlanningACCESSIBILITE(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFixe "ACCESSIBILITE", collabs, nb
End Sub
Sub GenererPlanningCMLEASING(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFixe "CM Leasing", collabs, nb
End Sub

' ============================================================
' GLF - VAGUES DE PAUSE + CONGÉS
' ============================================================
Sub GenererPlanningGLF(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GLF")
    EcrireEnTeteHorizontale ws, "GLF"

    Dim vaguesPause(1 To 5) As String
    vaguesPause(1) = "12:00": vaguesPause(2) = "12:30": vaguesPause(3) = "13:00"
    vaguesPause(4) = "13:30": vaguesPause(5) = "14:00"

    Dim glfIdx() As Integer
    Dim nbGLF As Integer
    nbGLF = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = "GLF" Then
            nbGLF = nbGLF + 1
            ReDim Preserve glfIdx(1 To nbGLF)
            glfIdx(nbGLF) = i
        End If
    Next i
    If nbGLF = 0 Then Exit Sub

    Dim ligne As Integer
    ligne = 4
    Dim k As Integer
    For k = 1 To nbGLF
        Dim idx As Integer
        idx = glfIdx(k)

        Dim groupeBase As Integer
        groupeBase = (k - 1) Mod 5
        Dim vagueIdx As Integer
        vagueIdx = ((groupeBase + collabs(idx).IndexRotation) Mod 5) + 1
        Dim pauseH As String: pauseH = vaguesPause(vagueIdx)
        Dim pauseF As String: pauseF = AjouterMinutes(pauseH, 60)

        Dim cellules(1 To 7) As String
        Dim entrees(1 To 7) As String
        Dim sorties(1 To 7) As String
        Dim pDs(1 To 7) As String
        Dim pFs(1 To 7) As String
        Dim j As Integer

        For j = 1 To 7
            Select Case j
                Case 1, 2, 3, 4
                    cellules(j) = FormatCelluleJour("08:00", "18:00", pauseH, pauseF)
                    entrees(j) = "08:00": sorties(j) = "18:00"
                    pDs(j) = pauseH: pFs(j) = pauseF
                Case 5
                    cellules(j) = FormatCelluleJour("08:00", "17:00", pauseH, pauseF)
                    entrees(j) = "08:00": sorties(j) = "17:00"
                    pDs(j) = pauseH: pFs(j) = pauseF
                Case Else
                    cellules(j) = "OFF"
                    entrees(j) = "": sorties(j) = "": pDs(j) = "": pFs(j) = ""
            End Select
        Next j

        AppliquerConges cellules, collabs(idx)
        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entrees, sorties, pDs, pFs
        ligne = ligne + 1
    Next k

    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "LÉGENDE VAGUES GLF (rotation par groupe de ~5, hebdomadaire)"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
    Dim v As Integer
    For v = 1 To 5
        ws.Cells(ligne + v, 1).Value = "Vague " & v & " : Pause " & vaguesPause(v) & "-" & AjouterMinutes(vaguesPause(v), 60)
    Next v
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' ============================================================
' EBRA - VAGUES DE PAUSE + CONGÉS
' ============================================================
Sub GenererPlanningEBRA(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("EBRA")
    EcrireEnTeteHorizontale ws, "EBRA"

    Dim vaguesPause(1 To 5) As String
    vaguesPause(1) = "11:00": vaguesPause(2) = "11:30": vaguesPause(3) = "12:00"
    vaguesPause(4) = "12:30": vaguesPause(5) = "13:00"

    Dim ebraIdx() As Integer
    Dim nbEBRA As Integer
    nbEBRA = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = "EBRA" Or UCase(Trim(collabs(i).projet)) = "EBRA PRESSE" Then
            nbEBRA = nbEBRA + 1
            ReDim Preserve ebraIdx(1 To nbEBRA)
            ebraIdx(nbEBRA) = i
        End If
    Next i
    If nbEBRA = 0 Then Exit Sub

    Dim ligne As Integer
    ligne = 4
    Dim k As Integer
    For k = 1 To nbEBRA
        Dim idx As Integer
        idx = ebraIdx(k)

        Dim groupeBase As Integer
        groupeBase = ((k - 1) \ 10) Mod 5
        Dim vagueIdx As Integer
        vagueIdx = ((groupeBase + collabs(idx).IndexRotation) Mod 5) + 1
        Dim pauseH As String: pauseH = vaguesPause(vagueIdx)
        Dim pauseF As String: pauseF = AjouterMinutes(pauseH, 60)

        Dim cellules(1 To 7) As String
        Dim entrees(1 To 7) As String
        Dim sorties(1 To 7) As String
        Dim pDs(1 To 7) As String
        Dim pFs(1 To 7) As String
        Dim j As Integer

        For j = 1 To 7
            Select Case j
                Case 1 To 5
                    cellules(j) = FormatCelluleJour("07:00", "16:00", pauseH, pauseF)
                    entrees(j) = "07:00": sorties(j) = "16:00"
                    pDs(j) = pauseH: pFs(j) = pauseF
                Case 6
                    cellules(j) = FormatCelluleJour("07:00", "11:00", "", "")
                    entrees(j) = "07:00": sorties(j) = "11:00"
                    pDs(j) = "": pFs(j) = ""
                Case 7
                    cellules(j) = "OFF"
                    entrees(j) = "": sorties(j) = "": pDs(j) = "": pFs(j) = ""
            End Select
        Next j

        AppliquerConges cellules, collabs(idx)
        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entrees, sorties, pDs, pFs
        ligne = ligne + 1
    Next k

    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "LÉGENDE VAGUES EBRA (rotation par groupe de ~10, hebdomadaire)"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
    Dim v As Integer
    For v = 1 To 5
        ws.Cells(ligne + v, 1).Value = "Vague " & v & " : Pause " & vaguesPause(v) & "-" & AjouterMinutes(vaguesPause(v), 60)
    Next v
    ws.Cells(ligne + 6, 1).Value = "Samedi : 07:00-11:00 sans pause | Dimanche : OFF"
    ws.Cells(ligne + 6, 1).Font.Italic = True
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' ============================================================
' GOOGLE LEADS - 7J/7 + SHIFTS ROTATIFS + ~5 OFF/JOUR
'
' Objectif : ~5 OFF par jour INCLUANT les congés
' Poids weekend renforcé : Dim et Sam ont un quota OFF plus élevé
'   → Dim: quotaMax = ceil(nbGL * 2/7) * 1.5
'   → Sam: quotaMax = ceil(nbGL * 2/7) * 1.3
'   → Jours semaine : quota normal
'
' ÉTAPE 1 : pré-comptage des congés par jour
' ÉTAPE 2 : attribution des OFF planifiés (complément jusqu'au quota ~5/j)
'           en commençant par Dim > Sam > ...
' ============================================================
Sub GenererPlanningGOOGLELEADS(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GOOGLE LEADS")
    EcrireEnTeteHorizontale ws, "GOOGLE LEADS"

    Dim entrees(1 To 5) As String
    Dim sorties(1 To 5) As String
    entrees(1) = "07:00": sorties(1) = "16:00"
    entrees(2) = "08:00": sorties(2) = "17:00"
    entrees(3) = "09:00": sorties(3) = "18:00"
    entrees(4) = "10:00": sorties(4) = "19:00"
    entrees(5) = "11:00": sorties(5) = "20:00"

    ' Collecter collabs GL
    Dim glIdx() As Integer
    Dim nbGL As Integer
    nbGL = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = "GOOGLE LEADS" Then
            nbGL = nbGL + 1
            ReDim Preserve glIdx(1 To nbGL)
            glIdx(nbGL) = i
        End If
    Next i
    If nbGL = 0 Then Exit Sub

    ' --- ÉTAPE 1 : pré-compter les congés par jour de la semaine ---
    Dim congeParJour(1 To 7) As Integer
    Dim j As Integer
    For j = 1 To 7
        congeParJour(j) = 0
    Next j
    Dim k As Integer
    For k = 1 To nbGL
        Dim idx As Integer
        idx = glIdx(k)
        For j = 1 To 7
            If EstEnConge(collabs(idx), DateDuJour(j)) Then
                congeParJour(j) = congeParJour(j) + 1
            End If
        Next j
    Next k

    ' --- ÉTAPE 2 : quotas OFF planifiés ---
    ' Objectif total OFF planifiés par jour = 5, mais les congés comptent dedans
    ' quotaOFF(j) = max(0, 5 - congeParJour(j))
    ' Pour Dim : quota = max(0, 7 - congeParJour(7))   ← plus de poids
    ' Pour Sam : quota = max(0, 6 - congeParJour(6))   ← poids intermédiaire
    Dim quotaOFF(1 To 7) As Integer
    For j = 1 To 7
        Select Case j
            Case 7  ' Dimanche
                quotaOFF(j) = Application.WorksheetFunction.Max(0, 7 - congeParJour(j))
            Case 6  ' Samedi
                quotaOFF(j) = Application.WorksheetFunction.Max(0, 6 - congeParJour(j))
            Case Else  ' Lun-Ven
                quotaOFF(j) = Application.WorksheetFunction.Max(0, 5 - congeParJour(j))
        End Select
    Next j

    ' Compteur OFF planifiés attribués (hors congé)
    Dim offPlanifieParJour(1 To 7) As Integer
    For j = 1 To 7
        offPlanifieParJour(j) = 0
    Next j

    Dim ligne As Integer
    ligne = 4

    For k = 1 To nbGL
        idx = glIdx(k)

        ' Shift selon rang + IndexRotation
        Dim shiftIdx As Integer
        shiftIdx = ((k - 1 + collabs(idx).IndexRotation) Mod 5) + 1

        Dim pD As String, pF As String
        pD = AjouterMinutes(entrees(shiftIdx), 300)
        pF = AjouterMinutes(pD, 60)

        ' Calcul des jours de congé de ce collab
        Dim joursConge(1 To 7) As Boolean
        For j = 1 To 7
            joursConge(j) = EstEnConge(collabs(idx), DateDuJour(j))
        Next j

        ' Calcul des 2 jours OFF planifiés (parmi les jours non-congé)
        ' → seulement si quota non atteint sur ce jour
        Dim off1 As Integer, off2 As Integer
        Call CalculerJoursOFF_GL_V5(offPlanifieParJour, quotaOFF, joursConge, off1, off2)

        ' Mettre à jour compteur
        If off1 > 0 Then offPlanifieParJour(off1) = offPlanifieParJour(off1) + 1
        If off2 > 0 Then offPlanifieParJour(off2) = offPlanifieParJour(off2) + 1

        Dim cellules(1 To 7) As String
        Dim entTab(1 To 7) As String
        Dim sorTab(1 To 7) As String
        Dim pdTab(1 To 7) As String
        Dim pfTab(1 To 7) As String

        For j = 1 To 7
            If joursConge(j) Then
                cellules(j) = "CONGE"
                entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            ElseIf j = off1 Or j = off2 Then
                cellules(j) = "OFF"
                entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            Else
                cellules(j) = FormatCelluleJour(entrees(shiftIdx), sorties(shiftIdx), pD, pF)
                entTab(j) = entrees(shiftIdx): sorTab(j) = sorties(shiftIdx)
                pdTab(j) = pD: pfTab(j) = pF
            End If
        Next j

        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entTab, sorTab, pdTab, pfTab
        ligne = ligne + 1
    Next k

    ' Légende
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "SHIFTS GOOGLE LEADS (rotation hebdomadaire par collab)"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
    Dim s As Integer
    For s = 1 To 5
        Dim pauseGL As String
        pauseGL = AjouterMinutes(entrees(s), 300)
        ws.Cells(ligne + s, 1).Value = "Shift " & s & " : " & entrees(s) & "-" & sorties(s) & _
                                        "  |  Pause: " & pauseGL & "-" & AjouterMinutes(pauseGL, 60)
    Next s
    ws.Cells(ligne + 6, 1).Value = "7j/7 | ~5 OFF/jour (congés inclus) | Poids renforcé Dim/Sam"
    ws.Cells(ligne + 6, 1).Font.Italic = True
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' Calcul des 2 jours OFF pour Google Leads V5
' - Respecte les quotas par jour
' - Ne place pas d'OFF sur un jour déjà en congé
' - Priorité : Dim > Sam > Ven > Jeu > Mer > Mar > Lun
' - Si quota déjà atteint sur un jour prioritaire, passe au suivant
Sub CalculerJoursOFF_GL_V5(offPlanifie() As Integer, quota() As Integer, _
                             joursConge() As Boolean, _
                             ByRef off1 As Integer, ByRef off2 As Integer)
    ' Ordre priorité : Dim=7, Sam=6, Ven=5, Jeu=4, Mer=3, Mar=2, Lun=1
    Dim priorite(1 To 7) As Integer
    priorite(1) = 7: priorite(2) = 6: priorite(3) = 5: priorite(4) = 4
    priorite(5) = 3: priorite(6) = 2: priorite(7) = 1

    Dim p As Integer
    off1 = 0: off2 = 0

    ' Premier OFF : jour prioritaire avec quota non atteint et pas en congé
    For p = 1 To 7
        Dim d As Integer: d = priorite(p)
        If Not joursConge(d) And offPlanifie(d) < quota(d) Then
            off1 = d
            Exit For
        End If
    Next p

    ' Deuxième OFF : idem, différent du premier
    For p = 1 To 7
        d = priorite(p)
        If d <> off1 And Not joursConge(d) And offPlanifie(d) < quota(d) Then
            off2 = d
            Exit For
        End If
    Next p

    ' Fallback : si quota dépassé partout, forcer sur les jours weekend
    If off1 = 0 Then off1 = 7
    If off2 = 0 Then
        For p = 1 To 7
            If priorite(p) <> off1 And Not joursConge(priorite(p)) Then
                off2 = priorite(p): Exit For
            End If
        Next p
    End If
    If off2 = 0 Then off2 = 6
End Sub

' ============================================================
' TLV - SHIFTS ROTATIFS + 2 JOURS OFF + CONGÉS
' ============================================================
Sub GenererPlanningTLV(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TLV")
    EcrireEnTeteHorizontale ws, "TLV"

    Dim entrees(1 To 2) As String
    Dim sorties(1 To 2) As String
    entrees(1) = "08:00": sorties(1) = "17:00"
    entrees(2) = "09:00": sorties(2) = "18:00"

    Dim joursReposSemaine(1 To 5) As Integer
    joursReposSemaine(1) = 1: joursReposSemaine(2) = 2: joursReposSemaine(3) = 3
    joursReposSemaine(4) = 4: joursReposSemaine(5) = 5

    ' Collecter collabs TLV
    Dim tlvIdx() As Integer
    Dim nbTLV As Integer
    nbTLV = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = "TLV" Or UCase(Trim(collabs(i).projet)) = "TELEVENTE" Then
            nbTLV = nbTLV + 1
            ReDim Preserve tlvIdx(1 To nbTLV)
            tlvIdx(nbTLV) = i
        End If
    Next i
    If nbTLV = 0 Then Exit Sub

    Dim ligne As Integer
    ligne = 4
    Dim k As Integer
    For k = 1 To nbTLV
        Dim idx As Integer
        idx = tlvIdx(k)

        ' Shift : rang + IndexRotation
        Dim shiftIdx As Integer
        shiftIdx = ((k - 1 + collabs(idx).IndexRotation) Mod 2) + 1

        Dim pD As String, pF As String
        pD = AjouterMinutes(entrees(shiftIdx), 300)
        pF = AjouterMinutes(pD, 60)

        ' Jour de repos semaine : distribué séquentiellement + rotation
        Dim reposBase As Integer: reposBase = (k - 1) Mod 5
        Dim reposRotated As Integer
        reposRotated = (reposBase + collabs(idx).IndexRotation) Mod 5
        Dim jourReposSem As Integer
        jourReposSem = joursReposSemaine(reposRotated + 1)

        Dim cellules(1 To 7) As String
        Dim entTab(1 To 7) As String
        Dim sorTab(1 To 7) As String
        Dim pdTab(1 To 7) As String
        Dim pfTab(1 To 7) As String
        Dim j As Integer

        For j = 1 To 7
            If j = 7 Or j = jourReposSem Then
                cellules(j) = "OFF"
                entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            Else
                cellules(j) = FormatCelluleJour(entrees(shiftIdx), sorties(shiftIdx), pD, pF)
                entTab(j) = entrees(shiftIdx): sorTab(j) = sorties(shiftIdx)
                pdTab(j) = pD: pfTab(j) = pF
            End If
        Next j

        ' Congés : remplace les jours travaillés en "CONGE"
        ' (les jours déjà OFF restent OFF, pas besoin de doubler)
        For j = 1 To 7
            If cellules(j) <> "OFF" And EstEnConge(collabs(idx), DateDuJour(j)) Then
                cellules(j) = "CONGE"
                entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            End If
        Next j

        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entTab, sorTab, pdTab, pfTab
        ligne = ligne + 1
    Next k

    If ligne > 4 Then
        ligne = ligne + 1
        ws.Cells(ligne, 1).Value = "SHIFTS TLV (rotation hebdomadaire)"
        ws.Cells(ligne, 1).Font.Bold = True
        ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
        ws.Cells(ligne + 1, 1).Value = "Shift 1 : 08:00-17:00 | Pause 13:00-14:00"
        ws.Cells(ligne + 2, 1).Value = "Shift 2 : 09:00-18:00 | Pause 14:00-15:00"
        ws.Cells(ligne + 3, 1).Value = "OFF : Dimanche fixe + 1 jour semaine rotatif (Lun-Ven), 1 par jour"
        ws.Cells(ligne + 3, 1).Font.Italic = True
    End If
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' ============================================================
' FACTO ET DAC - SHIFTS ROTATIFS + VENDREDI RÉDUIT + CONGÉS
' ============================================================
Sub GenererPlanningFactoDAC(nomFeuille As String, collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomFeuille)
    EcrireEnTeteHorizontale ws, nomFeuille

    Dim entrees(1 To 2) As String
    Dim sorties(1 To 2) As String
    entrees(1) = "07:00": sorties(1) = "17:00"
    entrees(2) = "08:00": sorties(2) = "18:00"

    ' Collecter collabs Facto/DAC avec rang pour rotation shift
    Dim fdIdx() As Integer
    Dim nbFD As Integer
    nbFD = 0
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = UCase(nomFeuille) Then
            nbFD = nbFD + 1
            ReDim Preserve fdIdx(1 To nbFD)
            fdIdx(nbFD) = i
        End If
    Next i
    If nbFD = 0 Then Exit Sub

    Dim ligne As Integer
    ligne = 4
    Dim k As Integer
    For k = 1 To nbFD
        Dim idx As Integer
        idx = fdIdx(k)

        ' Shift selon rang + IndexRotation (distribué dès sem 1)
        Dim shiftIdx As Integer
        shiftIdx = ((k - 1 + collabs(idx).IndexRotation) Mod 2) + 1

        Dim finNormale As String: finNormale = sorties(shiftIdx)
        Dim finVen As String: finVen = AjouterMinutes(finNormale, -60)
        Dim pD As String, pF As String
        pD = AjouterMinutes(entrees(shiftIdx), 300)
        pF = AjouterMinutes(pD, 60)

        Dim cellules(1 To 7) As String
        Dim entTab(1 To 7) As String
        Dim sorTab(1 To 7) As String
        Dim pdTab(1 To 7) As String
        Dim pfTab(1 To 7) As String
        Dim j As Integer

        For j = 1 To 7
            Select Case j
                Case 1, 2, 3, 4
                    cellules(j) = FormatCelluleJour(entrees(shiftIdx), finNormale, pD, pF)
                    entTab(j) = entrees(shiftIdx): sorTab(j) = finNormale
                    pdTab(j) = pD: pfTab(j) = pF
                Case 5
                    cellules(j) = FormatCelluleJour(entrees(shiftIdx), finVen, pD, pF)
                    entTab(j) = entrees(shiftIdx): sorTab(j) = finVen
                    pdTab(j) = pD: pfTab(j) = pF
                Case Else
                    cellules(j) = "OFF"
                    entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            End Select
        Next j

        ' Congés
        For j = 1 To 7
            If cellules(j) <> "OFF" And EstEnConge(collabs(idx), DateDuJour(j)) Then
                cellules(j) = "CONGE"
                entTab(j) = "": sorTab(j) = "": pdTab(j) = "": pfTab(j) = ""
            End If
        Next j

        EcrireLigneAvecConsolidation ws, ligne, collabs(idx), cellules, entTab, sorTab, pdTab, pfTab
        ligne = ligne + 1
    Next k

    If ligne > 4 Then
        ligne = ligne + 1
        ws.Cells(ligne, 1).Value = "SHIFTS " & UCase(nomFeuille) & " (rotation hebdomadaire)"
        ws.Cells(ligne, 1).Font.Bold = True
        ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
        ws.Cells(ligne + 1, 1).Value = "Shift 1 : 07:00-17:00 | Ven 07:00-16:00 | Pause 12:00-13:00"
        ws.Cells(ligne + 2, 1).Value = "Shift 2 : 08:00-18:00 | Ven 08:00-17:00 | Pause 13:00-14:00"
        ws.Cells(ligne + 3, 1).Value = "Sam-Dim : OFF"
        ws.Cells(ligne + 3, 1).Font.Italic = True
    End If
    AppliquerBorduresH ws, 4, ligne - 2
End Sub

Sub GenererPlanningFACTO(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFactoDAC "FACTO", collabs, nb
End Sub
Sub GenererPlanningDAC(collabs() As Collaborateur, nb As Integer)
    GenererPlanningFactoDAC "DAC", collabs, nb
End Sub

' ============================================================
' RESET ROTATIONS
' ============================================================
Sub ResetRotations()
    If MsgBox("Reinitialiser toutes les rotations ?", vbYesNo + vbWarning) = vbNo Then Exit Sub
    If FeuilleExiste("ROTATION") Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("ROTATION")
        If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
            ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, 5)).Clear
        End If
        MsgBox "Rotations reinitialisees.", vbInformation
    End If
End Sub
