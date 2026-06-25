' ============================================================
' GÉNÉRATEUR AUTOMATIQUE DE PLANNING HEBDOMADAIRE - V4
' FORMAT HORIZONTAL + ROTATIONS HORAIRES COMPLÈTES
' ============================================================
' RÈGLES PAR PROJET :
'
' FIXES (AFEDIM / ACCESSIBILITE / CM LEASING) :
'   Lun-Jeu 08:00-18:00 | Ven 08:00-17:00 | Sam-Dim OFF
'   Pause fixe 13:00-14:00 | Total = 44h/sem
'
' GLF :
'   Mêmes horaires que fixes
'   5 vagues de pause déjeuner par groupes (rotation hebdo) :
'   Vague 1 12:00 / Vague 2 12:30 / Vague 3 13:00 / Vague 4 13:30 / Vague 5 14:00
'   Chaque vague = groupe de ~5 personnes (réparti équitablement)
'
' EBRA :
'   Lun-Ven 07:00-16:00 | Sam 07:00-11:00 (sans pause) | Dim OFF
'   5 vagues de pause déjeuner par groupes de ~10 (rotation hebdo) :
'   Vague 1 11:00 / Vague 2 11:30 / Vague 3 12:00 / Vague 4 12:30 / Vague 5 13:00
'
' GOOGLE LEADS :
'   7j/7 avec 2 jours OFF/collab (objectif ~5 OFF/jour, priorité Sam/Dim)
'   Shifts rotatifs par collab (rotation hebdo sur 5 shifts) :
'     Shift 1 07:00-16:00 | Shift 2 08:00-17:00 | Shift 3 09:00-18:00
'     Shift 4 10:00-19:00 | Shift 5 11:00-20:00
'   Pause = heure entrée + 5h (durée 1h)
'
' TLV :
'   Lun-Sam | 2 jours OFF : Dim fixe + 1 jour semaine rotatif (Lun→Ven)
'   2 shifts rotatifs : 08:00-17:00 ou 09:00-18:00
'   Pause = entrée + 5h (durée 1h)
'
' FACTO / DAC :
'   Lun-Ven | Sam-Dim OFF
'   2 shifts rotatifs : 07:00-17:00 ou 08:00-18:00
'   Vendredi shift réduit : -1h en sortie
'   Pause = entrée + 5h (durée 1h)
' ============================================================

Option Explicit

Type Collaborateur
    NomComplet As String
    projet As String
    ville As String
    zone As String
    IndexRotation As Integer
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
    EffacerAnciensPlannings

    Dim collaborateurs() As Collaborateur
    Dim nbCollab As Integer
    nbCollab = LireCollaborateurs(collaborateurs)

    If nbCollab = 0 Then
        MsgBox "Aucun collaborateur trouve dans la feuille Utilisateurs."
        GoTo Cleanup
    End If

    GenererPlanningAFEDIM collaborateurs, nbCollab
    GenererPlanningACCESSIBILITE collaborateurs, nbCollab
    GenererPlanningCMLEASING collaborateurs, nbCollab
    GenererPlanningGLF collaborateurs, nbCollab
    GenererPlanningEBRA collaborateurs, nbCollab
    GenererPlanningGOOGLELEADS collaborateurs, nbCollab
    GenererPlanningTLV collaborateurs, nbCollab
    GenererPlanningFACTO collaborateurs, nbCollab
    GenererPlanningDAC collaborateurs, nbCollab
    MettreAJourRotation collaborateurs, nbCollab

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
    req = Split("Utilisateurs,AFEDIM,ACCESSIBILITE,CM Leasing,GLF,EBRA,GOOGLE LEADS,TLV,FACTO,DAC", ",")
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

' Ajouter des minutes à une heure "HH:MM"
Function AjouterMinutes(heure As String, minutes As Integer) As String
    If heure = "" Then AjouterMinutes = "": Exit Function
    Dim p() As String
    p = Split(heure, ":")
    Dim t As Integer
    t = CInt(p(0)) * 60 + CInt(p(1)) + minutes
    If t < 0 Then t = 0
    AjouterMinutes = Format(t \ 60, "00") & ":" & Format(t Mod 60, "00")
End Function

' Formater le contenu d'une cellule jour
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
' ÉCRIRE UNE LIGNE HORIZONTALE
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

        If cellules(j) = "OFF" Then
            cel.Interior.Color = RGB(255, 199, 206)
            cel.Font.Bold = True
            cel.Font.Color = RGB(192, 0, 0)
        Else
            cel.Font.Color = RGB(0, 0, 0)
            If ligne Mod 2 = 0 Then
                cel.Interior.Color = RGB(235, 241, 255)
            Else
                cel.Interior.Color = RGB(255, 255, 255)
            End If
        End If
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
        .Weight = xlMedium
        .Color = RGB(68, 114, 196)
    End With
    With rng.Borders(xlEdgeRight)
        .Weight = xlMedium
        .Color = RGB(68, 114, 196)
    End With
    With rng.Borders(xlEdgeTop)
        .Weight = xlMedium
        .Color = RGB(68, 114, 196)
    End With
    With rng.Borders(xlEdgeBottom)
        .Weight = xlMedium
        .Color = RGB(68, 114, 196)
    End With
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
End Sub

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
        collabs(i).NomComplet = Trim(ws.Cells(i + 1, 1).Value)
        collabs(i).projet = Trim(ws.Cells(i + 1, 2).Value)
        collabs(i).ville = Trim(ws.Cells(i + 1, 3).Value)
        collabs(i).zone = Trim(ws.Cells(i + 1, 4).Value)
        collabs(i).IndexRotation = LireIndexRotation(collabs(i).NomComplet, collabs(i).projet)
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
            Dim j As Integer
            For j = 1 To 7
                Select Case j
                    Case 1, 2, 3, 4  ' Lun-Jeu 08:00-18:00 pause 13:00-14:00
                        cellules(j) = FormatCelluleJour("08:00", "18:00", "13:00", "14:00")
                    Case 5           ' Ven 08:00-17:00 pause 13:00-14:00
                        cellules(j) = FormatCelluleJour("08:00", "17:00", "13:00", "14:00")
                    Case Else        ' Sam-Dim OFF
                        cellules(j) = "OFF"
                End Select
            Next j
            EcrireLigneHorizontale ws, ligne, collabs(i).NomComplet, collabs(i).ville, collabs(i).zone, cellules
            ligne = ligne + 1
        End If
    Next i

    ' Ajouter info totale heures
    If ligne > 4 Then
        ws.Cells(ligne + 1, 1).Value = "Total hebdomadaire : 44h (8h x 4j + 8h Ven + pauses 1h/j)"
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
' GLF - VAGUES DE PAUSE AVEC ROTATION HEBDO
' Mêmes horaires que fixes (08:00-18:00 LJ / 08:00-17:00 V)
' 5 vagues de pause : 12:00 / 12:30 / 13:00 / 13:30 / 14:00
' Chaque vague = groupe de ~5 personnes (répartition équitable)
' Rotation : chaque semaine les groupes avancent d'une vague
' ============================================================
Sub GenererPlanningGLF(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GLF")
    EcrireEnTeteHorizontale ws, "GLF"

    ' 5 vagues de pause (heure début, durée 1h)
    Dim vaguesPause(1 To 5) As String
    vaguesPause(1) = "12:00"
    vaguesPause(2) = "12:30"
    vaguesPause(3) = "13:00"
    vaguesPause(4) = "13:30"
    vaguesPause(5) = "14:00"

    ' Collecter les collabs GLF et trier par IndexRotation (pour grouper)
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

    ' Répartition équitable en 5 groupes
    ' Groupe 0 = vague 1, Groupe 1 = vague 2, etc.
    ' nbGLF = 21 → groupes de 5,5,5,3,3 par ex (floor(21/5)=4, 21 mod 5=1)
    ' La vague assignée = (groupeCollab + IndexRotation) mod 5 → rotation hebdo
    Dim ligne As Integer
    ligne = 4
    Dim k As Integer
    For k = 1 To nbGLF
        Dim idx As Integer
        idx = glfIdx(k)

        ' Groupe de base du collab (0 à 4), basé sur sa position dans l'équipe
        Dim groupeBase As Integer
        groupeBase = (k - 1) Mod 5  ' 0=vague1, 1=vague2, ..., 4=vague5

        ' Vague effective cette semaine = rotation selon IndexRotation
        Dim vagueIdx As Integer
        vagueIdx = ((groupeBase + collabs(idx).IndexRotation) Mod 5) + 1

        Dim pauseH As String
        Dim pauseF As String
        pauseH = vaguesPause(vagueIdx)
        pauseF = AjouterMinutes(pauseH, 60)

        Dim cellules(1 To 7) As String
        Dim j As Integer
        For j = 1 To 7
            Select Case j
                Case 1, 2, 3, 4  ' Lun-Jeu 08:00-18:00
                    cellules(j) = FormatCelluleJour("08:00", "18:00", pauseH, pauseF)
                Case 5           ' Ven 08:00-17:00
                    cellules(j) = FormatCelluleJour("08:00", "17:00", pauseH, pauseF)
                Case Else        ' Sam-Dim OFF
                    cellules(j) = "OFF"
            End Select
        Next j
        EcrireLigneHorizontale ws, ligne, collabs(idx).NomComplet, collabs(idx).ville, collabs(idx).zone, cellules
        ligne = ligne + 1
    Next k

    ' Légende vagues
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "LÉGENDE VAGUES GLF (rotation hebdomadaire par groupe de ~5)"
    ws.Cells(ligne, 1).Font.Bold = True
    ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
    Dim v As Integer
    For v = 1 To 5
        ws.Cells(ligne + v, 1).Value = "Vague " & v & " : Pause " & vaguesPause(v) & "-" & AjouterMinutes(vaguesPause(v), 60)
    Next v
    ws.Cells(ligne + 6, 1).Value = "Répartition : groupes de ~5 personnes, rotation +1 vague chaque semaine"
    ws.Cells(ligne + 6, 1).Font.Italic = True

    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' ============================================================
' EBRA - HORAIRES FIXES + VAGUES DE PAUSE AVEC ROTATION
' Lun-Ven 07:00-16:00 | Sam 07:00-11:00 (sans pause) | Dim OFF
' 5 vagues de pause par groupes de ~10 (rotation hebdo) :
'   Vague 1 11:00 / Vague 2 11:30 / Vague 3 12:00 / Vague 4 12:30 / Vague 5 13:00
' ============================================================
Sub GenererPlanningEBRA(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("EBRA")
    EcrireEnTeteHorizontale ws, "EBRA"

    Dim vaguesPause(1 To 5) As String
    vaguesPause(1) = "11:00"
    vaguesPause(2) = "11:30"
    vaguesPause(3) = "12:00"
    vaguesPause(4) = "12:30"
    vaguesPause(5) = "13:00"

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

    ' Répartition en groupes de ~10 personnes (5 vagues)
    ' Groupe de base = floor((k-1) / 10) mais on utilise mod 5 pour 5 vagues
    ' → chaque groupe successif de 10 collabs = 1 vague
    Dim ligne As Integer
    ligne = 4
    Dim k As Integer
    For k = 1 To nbEBRA
        Dim idx As Integer
        idx = ebraIdx(k)

        ' Groupe de base (0 à 4), groupes de 10 personnes
        Dim groupeBase As Integer
        groupeBase = ((k - 1) \ 10) Mod 5  ' groupes de 10, cycle sur 5 vagues

        ' Vague effective cette semaine
        Dim vagueIdx As Integer
        vagueIdx = ((groupeBase + collabs(idx).IndexRotation) Mod 5) + 1

        Dim pauseH As String
        Dim pauseF As String
        pauseH = vaguesPause(vagueIdx)
        pauseF = AjouterMinutes(pauseH, 60)

        Dim cellules(1 To 7) As String
        Dim j As Integer
        For j = 1 To 7
            Select Case j
                Case 1 To 5  ' Lun-Ven 07:00-16:00 avec pause
                    cellules(j) = FormatCelluleJour("07:00", "16:00", pauseH, pauseF)
                Case 6        ' Samedi 07:00-11:00 sans pause
                    cellules(j) = FormatCelluleJour("07:00", "11:00", "", "")
                Case 7        ' Dimanche OFF
                    cellules(j) = "OFF"
            End Select
        Next j
        EcrireLigneHorizontale ws, ligne, collabs(idx).NomComplet, collabs(idx).ville, collabs(idx).zone, cellules
        ligne = ligne + 1
    Next k

    ' Légende vagues
    ligne = ligne + 1
    ws.Cells(ligne, 1).Value = "LÉGENDE VAGUES EBRA (rotation hebdomadaire par groupe de ~10)"
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
' GOOGLE LEADS - 7J/7 + SHIFTS ROTATIFS + 2 JOURS OFF
'
' 5 shifts (rotation hebdo par collab via IndexRotation) :
'   Shift 1 : 07:00-16:00  Pause 12:00-13:00 (entrée + 5h)
'   Shift 2 : 08:00-17:00  Pause 13:00-14:00
'   Shift 3 : 09:00-18:00  Pause 14:00-15:00
'   Shift 4 : 10:00-19:00  Pause 15:00-16:00
'   Shift 5 : 11:00-20:00  Pause 16:00-17:00
'
' 2 jours OFF par collab | Objectif ~5 OFF/jour
' Priorité OFF : Dim > Sam > Ven > Jeu > Mer > Mar > Lun
' ============================================================
Sub GenererPlanningGOOGLELEADS(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("GOOGLE LEADS")
    EcrireEnTeteHorizontale ws, "GOOGLE LEADS"

    ' 5 shifts : entrée, sortie (entrée + 9h)
    Dim entrees(1 To 5) As String
    Dim sorties(1 To 5) As String
    entrees(1) = "07:00": sorties(1) = "16:00"
    entrees(2) = "08:00": sorties(2) = "17:00"
    entrees(3) = "09:00": sorties(3) = "18:00"
    entrees(4) = "10:00": sorties(4) = "19:00"
    entrees(5) = "11:00": sorties(5) = "20:00"
    ' Pause = entrée + 5h (durée 1h)

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

    ' Compteur OFF par jour pour équilibrage
    Dim offParJour(1 To 7) As Integer
    Dim j As Integer
    For j = 1 To 7
        offParJour(j) = 0
    Next j

    Dim ligne As Integer
    ligne = 4
    Dim k As Integer
    For k = 1 To nbGL
        Dim idx As Integer
        idx = glIdx(k)

        ' Shift : basé sur rang dans l'équipe (k) + IndexRotation pour rotation hebdo
        ' → dès la semaine 1 : collab 1=shift1, collab 2=shift2, ..., collab 6=shift1, etc.
        ' → semaine suivante : tout le monde avance d'un shift (IndexRotation+1)
        Dim shiftIdx As Integer
        shiftIdx = ((k - 1 + collabs(idx).IndexRotation) Mod 5) + 1

        ' Pause = entrée + 5h
        Dim pD As String, pF As String
        pD = AjouterMinutes(entrees(shiftIdx), 300)
        pF = AjouterMinutes(pD, 60)

        ' Calcul des 2 jours OFF (équilibré, priorité Dim > Sam > ...)
        Dim off1 As Integer, off2 As Integer
        Call CalculerJoursOFF_GL(offParJour, off1, off2)
        offParJour(off1) = offParJour(off1) + 1
        offParJour(off2) = offParJour(off2) + 1

        Dim cellules(1 To 7) As String
        For j = 1 To 7
            If j = off1 Or j = off2 Then
                cellules(j) = "OFF"
            Else
                cellules(j) = FormatCelluleJour(entrees(shiftIdx), sorties(shiftIdx), pD, pF)
            End If
        Next j

        EcrireLigneHorizontale ws, ligne, collabs(idx).NomComplet, collabs(idx).ville, collabs(idx).zone, cellules
        ligne = ligne + 1
    Next k

    ' Légende shifts
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
    ws.Cells(ligne + 6, 1).Value = "7j/7 | 2 jours OFF/collab | ~5 OFF/jour | Priorité OFF : Dim > Sam"
    ws.Cells(ligne + 6, 1).Font.Italic = True

    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' Calcul des 2 jours OFF avec priorité Dim > Sam > Ven > Jeu > Mer > Mar > Lun
' Équilibrage par compteur offParJour
Sub CalculerJoursOFF_GL(offParJour() As Integer, ByRef off1 As Integer, ByRef off2 As Integer)
    ' Ordre priorité : 7(Dim), 6(Sam), 5(Ven), 4(Jeu), 3(Mer), 2(Mar), 1(Lun)
    Dim priorite(1 To 7) As Integer
    priorite(1) = 7
    priorite(2) = 6
    priorite(3) = 5
    priorite(4) = 4
    priorite(5) = 3
    priorite(6) = 2
    priorite(7) = 1

    Dim minOff As Integer
    Dim bestJ As Integer
    Dim p As Integer

    ' Premier jour OFF : le moins chargé parmi les plus prioritaires
    minOff = 9999
    bestJ = 7
    For p = 1 To 7
        If offParJour(priorite(p)) < minOff Then
            minOff = offParJour(priorite(p))
            bestJ = priorite(p)
        End If
    Next p
    off1 = bestJ

    ' Deuxième jour OFF : le moins chargé parmi les restants
    minOff = 9999
    bestJ = 6
    For p = 1 To 7
        If priorite(p) <> off1 And offParJour(priorite(p)) < minOff Then
            minOff = offParJour(priorite(p))
            bestJ = priorite(p)
        End If
    Next p
    off2 = bestJ
End Sub

' ============================================================
' TLV - SHIFTS ROTATIFS + 2 JOURS OFF (Dim fixe + 1j semaine)
'
' 2 shifts rotatifs (rotation hebdo via IndexRotation Mod 2) :
'   Shift 1 : 08:00-17:00
'   Shift 2 : 09:00-18:00
' Pause = entrée + 5h (durée 1h)
' Jours travaillés : Lun-Sam | Dim OFF fixe
' Repos semaine rotatif : 1 jour parmi Lun-Ven (rotation sur 5 positions)
' ============================================================
Sub GenererPlanningTLV(collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TLV")
    EcrireEnTeteHorizontale ws, "TLV"

    Dim entrees(1 To 2) As String
    Dim sorties(1 To 2) As String
    entrees(1) = "08:00": sorties(1) = "17:00"
    entrees(2) = "09:00": sorties(2) = "18:00"

    ' Jours de repos semaine (Lun à Ven seulement, Sam travaillé)
    Dim joursReposSemaine(1 To 5) As Integer
    joursReposSemaine(1) = 1  ' Lundi
    joursReposSemaine(2) = 2  ' Mardi
    joursReposSemaine(3) = 3  ' Mercredi
    joursReposSemaine(4) = 4  ' Jeudi
    joursReposSemaine(5) = 5  ' Vendredi

    ' Collecter d'abord les collabs TLV pour connaître leur rang dans l'équipe
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

        ' Shift : basé sur rang dans l'équipe + IndexRotation (rotation hebdo)
        ' → les collabs alternent shift 1/2 selon leur rang, et ça tourne chaque semaine
        Dim shiftIdx As Integer
        shiftIdx = ((k - 1 + collabs(idx).IndexRotation) Mod 2) + 1

        ' Pause = entrée + 5h
        Dim pD As String, pF As String
        pD = AjouterMinutes(entrees(shiftIdx), 300)
        pF = AjouterMinutes(pD, 60)

        ' Jour de repos en semaine :
        ' → distribué séquentiellement par rang (1=Lun, 2=Mar, ..., 5=Ven, 6=Lun, ...)
        ' → puis décalé par IndexRotation pour la rotation hebdo
        ' → garantit que chaque jour accueille au plus ceil(nbTLV/5) personnes en repos
        Dim reposBase As Integer
        reposBase = (k - 1) Mod 5  ' rang de base 0..4

        Dim reposRotated As Integer
        reposRotated = (reposBase + collabs(idx).IndexRotation) Mod 5  ' rotation hebdo

        Dim jourReposSem As Integer
        jourReposSem = joursReposSemaine(reposRotated + 1)

        Dim cellules(1 To 7) As String
        Dim j As Integer
        For j = 1 To 7
            If j = 7 Then                    ' Dimanche = OFF fixe
                cellules(j) = "OFF"
            ElseIf j = jourReposSem Then     ' Repos semaine (1 par jour, réparti)
                cellules(j) = "OFF"
            Else                             ' Jours travaillés
                cellules(j) = FormatCelluleJour(entrees(shiftIdx), sorties(shiftIdx), pD, pF)
            End If
        Next j
        EcrireLigneHorizontale ws, ligne, collabs(idx).NomComplet, collabs(idx).ville, collabs(idx).zone, cellules
        ligne = ligne + 1
    Next k

    ' Légende
    If ligne > 4 Then
        ligne = ligne + 1
        ws.Cells(ligne, 1).Value = "SHIFTS TLV (rotation hebdomadaire)"
        ws.Cells(ligne, 1).Font.Bold = True
        ws.Cells(ligne, 1).Font.Color = RGB(31, 73, 125)
        ws.Cells(ligne + 1, 1).Value = "Shift 1 : 08:00-17:00 | Pause 13:00-14:00"
        ws.Cells(ligne + 2, 1).Value = "Shift 2 : 09:00-18:00 | Pause 14:00-15:00"
        ws.Cells(ligne + 3, 1).Value = "OFF : Dimanche fixe + 1 jour semaine rotatif (Lun→Ven)"
        ws.Cells(ligne + 3, 1).Font.Italic = True
    End If

    AppliquerBorduresH ws, 4, ligne - 2
End Sub

' ============================================================
' FACTO ET DAC - SHIFTS ROTATIFS + VENDREDI RÉDUIT
'
' 2 shifts rotatifs (rotation hebdo via IndexRotation Mod 2) :
'   Shift 1 : 07:00-17:00 | Ven 07:00-16:00
'   Shift 2 : 08:00-18:00 | Ven 08:00-17:00
' Pause = entrée + 5h (durée 1h)
' Lun-Ven | Sam-Dim OFF
' ============================================================
Sub GenererPlanningFactoDAC(nomFeuille As String, collabs() As Collaborateur, nb As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomFeuille)
    EcrireEnTeteHorizontale ws, nomFeuille

    Dim entrees(1 To 2) As String
    Dim sorties(1 To 2) As String
    entrees(1) = "07:00": sorties(1) = "17:00"
    entrees(2) = "08:00": sorties(2) = "18:00"

    Dim ligne As Integer
    ligne = 4
    Dim i As Integer
    For i = 1 To nb
        If UCase(Trim(collabs(i).projet)) = UCase(nomFeuille) Then
            ' Shift selon rotation hebdo (2 shifts)
            Dim shiftIdx As Integer
            shiftIdx = (collabs(i).IndexRotation Mod 2) + 1

            ' Vendredi : sortie -1h
            Dim finNormale As String
            Dim finVen As String
            finNormale = sorties(shiftIdx)
            finVen = AjouterMinutes(finNormale, -60)

            ' Pause = entrée + 5h (même toute la semaine, incl. vendredi)
            Dim pD As String, pF As String
            pD = AjouterMinutes(entrees(shiftIdx), 300)
            pF = AjouterMinutes(pD, 60)

            Dim cellules(1 To 7) As String
            Dim j As Integer
            For j = 1 To 7
                Select Case j
                    Case 1, 2, 3, 4  ' Lun-Jeu : horaire normal
                        cellules(j) = FormatCelluleJour(entrees(shiftIdx), finNormale, pD, pF)
                    Case 5           ' Vendredi : sortie avancée de 1h
                        cellules(j) = FormatCelluleJour(entrees(shiftIdx), finVen, pD, pF)
                    Case Else        ' Sam-Dim OFF
                        cellules(j) = "OFF"
                End Select
            Next j
            EcrireLigneHorizontale ws, ligne, collabs(i).NomComplet, collabs(i).ville, collabs(i).zone, cellules
            ligne = ligne + 1
        End If
    Next i

    ' Légende
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
