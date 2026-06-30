' ============================================================
' UFParametre — UserForm : Modification d'un créneau de planning
' ============================================================
' UTILISATION :
'   1. Dans le VBE, insérer un UserForm nommé "UFParametre"
'   2. Coller ce code dans son module Code
'   3. Créer les contrôles décrits dans la section CONTROLES ci-dessous
'
' CONTROLES A CREER (noms exacts) :
'   Labels      : lblTitre, lblProjet, lblCollab, lblJour, lblInfo
'                 lblEntree, lblSortie, lblPauseD, lblPauseF
'   ComboBox    : cboProjet, cboCollab, cboJour
'   TextBox     : txtEntree, txtSortie, txtPauseD, txtPauseF
'   CheckBox    : chkOFF, chkCONGE, chkTT, chkRENFORT, chkSHIFT
'   Frame       : frmHoraires (contient les 4 TextBox horaires + labels)
'   CommandButton: btnValider, btnAnnuler, btnReset
' ============================================================

Option Explicit

' ──────────────────────────────────────────────
' VARIABLES MODULE
' ──────────────────────────────────────────────
Private m_projet    As String
Private m_collab    As String
Private m_jourIdx   As Integer   ' 1=Lun … 7=Dim
Private m_colJour   As Integer   ' colonne dans feuille projet (4 à 10)
Private m_ligneCollab As Long    ' ligne du collab dans feuille projet
Private m_wsProjFeuille As Worksheet
Private m_wsPlanning As Worksheet
Private m_wsConsol  As Worksheet
Private m_projets() As String
Private m_jours(1 To 7) As String

' ──────────────────────────────────────────────
' INITIALISATION
' ──────────────────────────────────────────────
Private Sub UserForm_Initialize()
    ' Titre
    Me.Caption = "UFParametre — Modifier un créneau"
    lblTitre.Caption = "Modifier un créneau de planning"

    ' Jours
    m_jours(1) = "Lundi"
    m_jours(2) = "Mardi"
    m_jours(3) = "Mercredi"
    m_jours(4) = "Jeudi"
    m_jours(5) = "Vendredi"
    m_jours(6) = "Samedi"
    m_jours(7) = "Dimanche"

    ' Remplir liste Projets (feuilles projet connues)
    m_projets = Split("AFEDIM,ACCESSIBILITE,CM Leasing,GLF,EBRA,GOOGLE LEADS,TLV,FACTO,DAC", ",")
    cboProjet.Clear
    Dim p As Integer
    For p = 0 To UBound(m_projets)
        If FeuilleExiste(m_projets(p)) Then
            cboProjet.AddItem m_projets(p)
        End If
    Next p

    ' Remplir liste Jours
    cboJour.Clear
    Dim j As Integer
    For j = 1 To 7
        cboJour.AddItem m_jours(j)
    Next j

    ' Etat initial : champs horaires désactivés
    ViderFormulaire
    frmHoraires.Enabled = False
    btnValider.Enabled = False

    lblInfo.Caption = "Sélectionnez un projet, un collaborateur et un jour."
End Sub

' ──────────────────────────────────────────────
' SELECTION PROJET → charger collaborateurs
' ──────────────────────────────────────────────
Private Sub cboProjet_Change()
    m_projet = Trim(cboProjet.Value)
    cboCollab.Clear
    cboJour.ListIndex = -1
    ViderFormulaire
    frmHoraires.Enabled = False
    btnValider.Enabled = False
    lblInfo.Caption = "Sélectionnez un collaborateur."

    If m_projet = "" Then Exit Sub
    If Not FeuilleExiste(m_projet) Then
        lblInfo.Caption = "Feuille introuvable : " & m_projet
        Exit Sub
    End If

    Set m_wsProjFeuille = ThisWorkbook.Sheets(m_projet)

    ' Les collaborateurs sont en col A à partir de la ligne 4
    Dim lr As Long
    lr = m_wsProjFeuille.Cells(m_wsProjFeuille.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 4 To lr
        Dim v As String
        v = Trim(CStr(m_wsProjFeuille.Cells(i, 1).Value))
        If v <> "" Then cboCollab.AddItem v
    Next i

    If cboCollab.ListCount = 0 Then
        lblInfo.Caption = "Aucun collaborateur trouvé dans la feuille " & m_projet & "."
    End If
End Sub

' ──────────────────────────────────────────────
' SELECTION COLLABORATEUR
' ──────────────────────────────────────────────
Private Sub cboCollab_Change()
    m_collab = Trim(cboCollab.Value)
    cboJour.ListIndex = -1
    ViderFormulaire
    frmHoraires.Enabled = False
    btnValider.Enabled = False

    If m_collab = "" Then Exit Sub
    lblInfo.Caption = "Sélectionnez un jour de la semaine."
End Sub

' ──────────────────────────────────────────────
' SELECTION JOUR → charger les données actuelles
' ──────────────────────────────────────────────
Private Sub cboJour_Change()
    If cboJour.ListIndex < 0 Then Exit Sub
    m_jourIdx = cboJour.ListIndex + 1   ' 1=Lundi … 7=Dimanche
    m_colJour = 3 + m_jourIdx           ' Col 4=Lun … Col 10=Dim

    If m_projet = "" Or m_collab = "" Then Exit Sub

    ' Trouver la ligne du collaborateur
    m_ligneCollab = TrouverLigneCollab(m_wsProjFeuille, m_collab)
    If m_ligneCollab = 0 Then
        lblInfo.Caption = "Collaborateur introuvable dans la feuille."
        Exit Sub
    End If

    ' Charger les données actuelles dans le formulaire
    ChargerDonneesCreneau

    frmHoraires.Enabled = True
    btnValider.Enabled = True
    lblInfo.Caption = "Données chargées. Modifiez puis cliquez Valider."
End Sub

' ──────────────────────────────────────────────
' CHARGER DONNEES DEPUIS FEUILLE PROJET
' ──────────────────────────────────────────────
Private Sub ChargerDonneesCreneau()
    Dim cel As String
    cel = Trim(CStr(m_wsProjFeuille.Cells(m_ligneCollab, m_colJour).Value))

    ' Réinitialiser tous les contrôles
    ViderFormulaire

    ' Analyser le contenu de la cellule
    Dim lignes() As String
    lignes = Split(cel, Chr(10))

    Dim estOFF As Boolean, estCONGE As Boolean
    Dim estTT As Boolean, estRENFORT As Boolean, estSHIFT As Boolean
    Dim horaire As String, pauseStr As String

    estOFF = (UCase(Trim(cel)) = "OFF" Or cel = "")
    estCONGE = (UCase(Trim(cel)) = "CONGE")
    estTT = (Left(UCase(cel), 2) = "TT")
    estRENFORT = (InStr(cel, "[RENFORT]") > 0)
    estSHIFT = (InStr(cel, "[SHIFT R") > 0)

    chkOFF.Value = estOFF
    chkCONGE.Value = estCONGE
    chkTT.Value = estTT
    chkRENFORT.Value = estRENFORT
    chkSHIFT.Value = estSHIFT

    If estOFF Or estCONGE Then
        ' Pas d'horaires à afficher
        txtEntree.Text = ""
        txtSortie.Text = ""
        txtPauseD.Text = ""
        txtPauseF.Text = ""
    Else
        ' Extraire la première ligne d'horaire (retire le préfixe "TT " si présent)
        Dim l0 As String
        l0 = Trim(lignes(0))
        If Left(UCase(l0), 3) = "TT " Then l0 = Trim(Mid(l0, 4))

        ' Format attendu : "HH:MM - HH:MM"
        Dim parts() As String
        parts = Split(l0, " - ")
        If UBound(parts) >= 1 Then
            txtEntree.Text = Trim(parts(0))
            txtSortie.Text = Trim(parts(1))
        End If

        ' Chercher la ligne Pause si elle existe
        Dim li As Integer
        For li = 1 To UBound(lignes)
            Dim lx As String: lx = Trim(lignes(li))
            If Left(UCase(lx), 5) = "PAUSE" Then
                ' Format : "Pause: HH:MM-HH:MM"
                Dim pStr As String
                pStr = Trim(Mid(lx, InStr(lx, ":") + 1))
                Dim pp() As String
                pp = Split(pStr, "-")
                If UBound(pp) >= 1 Then
                    txtPauseD.Text = Trim(pp(0))
                    txtPauseF.Text = Trim(pp(1))
                End If
                Exit For
            End If
        Next li
    End If

    ' Afficher un résumé
    lblInfo.Caption = "Créneau actuel : " & IIf(cel = "", "(vide)", cel)
End Sub

' ──────────────────────────────────────────────
' CHECKBOX : gestion exclusive OFF / CONGE
' ──────────────────────────────────────────────
Private Sub chkOFF_Click()
    If chkOFF.Value Then
        chkCONGE.Value = False
        txtEntree.Text = ""
        txtSortie.Text = ""
        txtPauseD.Text = ""
        txtPauseF.Text = ""
        txtEntree.Enabled = False
        txtSortie.Enabled = False
        txtPauseD.Enabled = False
        txtPauseF.Enabled = False
    Else
        txtEntree.Enabled = True
        txtSortie.Enabled = True
        txtPauseD.Enabled = True
        txtPauseF.Enabled = True
    End If
End Sub

Private Sub chkCONGE_Click()
    If chkCONGE.Value Then
        chkOFF.Value = False
        txtEntree.Text = ""
        txtSortie.Text = ""
        txtPauseD.Text = ""
        txtPauseF.Text = ""
        txtEntree.Enabled = False
        txtSortie.Enabled = False
        txtPauseD.Enabled = False
        txtPauseF.Enabled = False
    Else
        txtEntree.Enabled = True
        txtSortie.Enabled = True
        txtPauseD.Enabled = True
        txtPauseF.Enabled = True
    End If
End Sub

' ──────────────────────────────────────────────
' BOUTON VALIDER
' ──────────────────────────────────────────────
Private Sub btnValider_Click()
    ' ── Validation saisie ──
    If Not ValiderSaisie() Then Exit Sub

    ' ── Construire la nouvelle valeur cellule ──
    Dim nouvelleValeur As String
    nouvelleValeur = ConstruireValeurCellule()

    ' ── Construire entrée / sortie brutes (sans préfixes) ──
    Dim entree As String, sortie As String
    Dim pauseD As String, pauseF As String
    Dim activite As String

    If chkOFF.Value Then
        entree = "": sortie = "": pauseD = "": pauseF = ""
        activite = "OFF"
    ElseIf chkCONGE.Value Then
        entree = "": sortie = "": pauseD = "": pauseF = ""
        activite = "CONGE"
    Else
        entree = Trim(txtEntree.Text)
        sortie = Trim(txtSortie.Text)
        pauseD = Trim(txtPauseD.Text)
        pauseF = Trim(txtPauseF.Text)
        If chkTT.Value Then
            activite = "TT"
        ElseIf chkRENFORT.Value Then
            activite = "RENFORT"
        Else
            activite = m_projet
        End If
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrValider

    ' ── 1. Mettre à jour la feuille PROJET (tableau horizontal) ──
    MAJ_FeuilleProjHorizontale nouvelleValeur

    ' ── 2. Mettre à jour la feuille PLANNING (ligne par collaborateur) ──
    Set m_wsPlanning = ThisWorkbook.Sheets("PLANNING")
    MAJ_FeuillePlanning entree, sortie, activite

    ' ── 3. Mettre à jour la feuille CONSOLIDATION ──
    Set m_wsConsol = ThisWorkbook.Sheets("CONSOLIDATION")
    MAJ_Consolidation entree, sortie, pauseD, pauseF, activite

    ' ── 4. Recalculer les cumuls ──
    CalculerCumulsSemaine

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    lblInfo.Caption = "✅ Mise à jour effectuée avec succès."
    MsgBox "Créneau mis à jour !" & Chr(10) & _
           m_collab & " — " & m_jours(m_jourIdx) & Chr(10) & _
           "Nouvelle valeur : " & nouvelleValeur, vbInformation, "UFParametre"
    Exit Sub

ErrValider:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Erreur " & Err.Number & " : " & Err.Description, vbCritical, "UFParametre"
End Sub

' ──────────────────────────────────────────────
' CONSTRUIRE LA VALEUR DE CELLULE (feuille projet)
' ──────────────────────────────────────────────
Private Function ConstruireValeurCellule() As String
    Dim v As String

    If chkOFF.Value Then
        v = "OFF"
    ElseIf chkCONGE.Value Then
        v = "CONGE"
    Else
        Dim ent As String: ent = Trim(txtEntree.Text)
        Dim sor As String: sor = Trim(txtSortie.Text)
        Dim pD  As String: pD  = Trim(txtPauseD.Text)
        Dim pF  As String: pF  = Trim(txtPauseF.Text)

        ' Ligne horaire principale
        v = ent & " - " & sor

        ' Ligne pause si renseignée
        If pD <> "" And pF <> "" Then
            v = v & Chr(10) & "Pause: " & pD & "-" & pF
        End If

        ' Tags optionnels
        If chkSHIFT.Value Then v = v & Chr(10) & "[SHIFT Réduit]"
        If chkRENFORT.Value Then v = v & Chr(10) & "[RENFORT]"

        ' Préfixe TT
        If chkTT.Value Then v = "TT " & v
    End If

    ConstruireValeurCellule = v
End Function

' ──────────────────────────────────────────────
' MISE A JOUR — FEUILLE PROJET (tableau horizontal)
' ──────────────────────────────────────────────
Private Sub MAJ_FeuilleProjHorizontale(nouvelleValeur As String)
    Dim cel As Range
    Set cel = m_wsProjFeuille.Cells(m_ligneCollab, m_colJour)
    cel.Value = nouvelleValeur
    cel.WrapText = True
    cel.HorizontalAlignment = xlCenter
    cel.VerticalAlignment = xlCenter

    ' Couleur selon statut
    AppliquerCouleurCellule cel, nouvelleValeur
End Sub

' ──────────────────────────────────────────────
' MISE A JOUR — FEUILLE PLANNING
' Structure : 1 ligne par collab par semaine
' Colonnes entrée/sortie : Lun=12/13, Mar=14/15 … Dim=24/25
' ──────────────────────────────────────────────
Private Sub MAJ_FeuillePlanning(entree As String, sortie As String, activite As String)
    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(LundiSemaine(), 2)

    Dim colE As Integer: colE = 10 + (m_jourIdx * 2)
    Dim colS As Integer: colS = colE + 1

    Dim lr As Long
    Dim lastRow As Long
    lastRow = m_wsPlanning.Cells(m_wsPlanning.Rows.Count, 1).End(xlUp).Row
    lr = 0
    Dim i As Long
    For i = 2 To lastRow
        If CStr(m_wsPlanning.Cells(i, 1).Value) = CStr(sem) And _
           m_wsPlanning.Cells(i, 5).Value = m_collab Then
            lr = i: Exit For
        End If
    Next i

    If lr = 0 Then
        ' Ligne non trouvée : on ne peut pas créer sans toutes les infos du collab
        ' → log silencieux (le planning sera régénéré au prochain GenererPlanning)
        Exit Sub
    End If

    Dim valE As String, valS As String
    Select Case True
        Case activite = "CONGE"
            valE = "CONGE": valS = "CONGE"
        Case activite = "OFF" Or entree = ""
            valE = "OFF": valS = "OFF"
        Case activite = "TT"
            valE = "TT " & entree: valS = "TT " & sortie
        Case Else
            valE = entree: valS = sortie
    End Select

    Dim celE As Range: Set celE = m_wsPlanning.Cells(lr, colE)
    Dim celS As Range: Set celS = m_wsPlanning.Cells(lr, colS)
    celE.Value = valE: celS.Value = valS
    celE.HorizontalAlignment = xlCenter: celS.HorizontalAlignment = xlCenter

    ' Couleurs
    Select Case True
        Case valE = "OFF"
            celE.Interior.Color = RGB(255, 199, 206): celE.Font.Color = RGB(192, 0, 0): celE.Font.Bold = True
            celS.Interior.Color = RGB(255, 199, 206): celS.Font.Color = RGB(192, 0, 0): celS.Font.Bold = True
        Case valE = "CONGE"
            celE.Interior.Color = RGB(255, 230, 153): celE.Font.Color = RGB(156, 87, 0): celE.Font.Bold = True
            celS.Interior.Color = RGB(255, 230, 153): celS.Font.Color = RGB(156, 87, 0): celS.Font.Bold = True
        Case Left(valE, 2) = "TT"
            celE.Interior.Color = RGB(220, 190, 255): celE.Font.Color = RGB(70, 0, 130): celE.Font.Bold = False
            celS.Interior.Color = RGB(220, 190, 255): celS.Font.Color = RGB(70, 0, 130): celS.Font.Bold = False
        Case Else
            celE.Interior.ColorIndex = xlNone: celE.Font.Color = RGB(0, 0, 0): celE.Font.Bold = False
            celS.Interior.ColorIndex = xlNone: celS.Font.Color = RGB(0, 0, 0): celS.Font.Bold = False
    End Select
End Sub

' ──────────────────────────────────────────────
' MISE A JOUR — FEUILLE CONSOLIDATION
' ──────────────────────────────────────────────
Private Sub MAJ_Consolidation(entree As String, sortie As String, _
                               pauseD As String, pauseF As String, activite As String)
    Dim dateCreneau As Date
    dateCreneau = LundiSemaine() + (m_jourIdx - 1)
    Dim sem As Integer
    sem = Application.WorksheetFunction.WeekNum(dateCreneau, 2)

    ' Chercher la ligne existante (collab + date)
    Dim lr As Long
    Dim lastRow As Long
    lastRow = m_wsConsol.Cells(m_wsConsol.Rows.Count, 1).End(xlUp).Row
    lr = 0
    Dim i As Long
    For i = 2 To lastRow
        If m_wsConsol.Cells(i, 1).Value = m_collab Then
            Dim dCell As Variant: dCell = m_wsConsol.Cells(i, 2).Value
            If IsDate(dCell) Then
                If CDate(dCell) = dateCreneau Then
                    lr = i: Exit For
                End If
            End If
        End If
    Next i

    If lr = 0 Then
        ' Ligne absente : on ne peut pas créer sans le collab complet
        Exit Sub
    End If

    ' Mettre à jour entrée, sortie, pause, activité
    m_wsConsol.Cells(lr, 3).Value = IIf(activite = "OFF" Or activite = "CONGE", "", entree)
    m_wsConsol.Cells(lr, 4).Value = IIf(activite = "OFF" Or activite = "CONGE", "", sortie)
    m_wsConsol.Cells(lr, 5).Value = IIf(activite = "OFF" Or activite = "CONGE", "", pauseD)
    m_wsConsol.Cells(lr, 6).Value = IIf(activite = "OFF" Or activite = "CONGE", "", pauseF)
    m_wsConsol.Cells(lr, 8).Value = activite

    ' Couleur de ligne
    Select Case activite
        Case "OFF":   m_wsConsol.Rows(lr).Interior.Color = RGB(255, 199, 206)
        Case "CONGE": m_wsConsol.Rows(lr).Interior.Color = RGB(255, 230, 153)
        Case "TT":    m_wsConsol.Rows(lr).Interior.Color = RGB(230, 210, 255)
        Case Else
            If lr Mod 2 = 0 Then
                m_wsConsol.Rows(lr).Interior.Color = RGB(235, 241, 255)
            Else
                m_wsConsol.Rows(lr).Interior.Color = RGB(255, 255, 255)
            End If
    End Select
End Sub

' ──────────────────────────────────────────────
' APPLIQUER COULEUR SUR CELLULE (feuille projet)
' ──────────────────────────────────────────────
Private Sub AppliquerCouleurCellule(cel As Range, valeur As String)
    Select Case True
        Case UCase(valeur) = "OFF"
            cel.Interior.Color = RGB(255, 199, 206)
            cel.Font.Bold = True
            cel.Font.Color = RGB(192, 0, 0)
        Case UCase(valeur) = "CONGE"
            cel.Interior.Color = RGB(255, 192, 0)
            cel.Font.Bold = True
            cel.Font.Color = RGB(0, 0, 0)
        Case Left(UCase(valeur), 2) = "TT"
            cel.Interior.Color = RGB(220, 190, 255)
            cel.Font.Bold = False
            cel.Font.Color = RGB(70, 0, 130)
        Case InStr(valeur, "[RENFORT]") > 0
            cel.Interior.Color = RGB(169, 208, 142)
            cel.Font.Bold = True
            cel.Font.Color = RGB(0, 97, 0)
        Case InStr(valeur, "[SHIFT R") > 0
            cel.Interior.Color = RGB(255, 220, 140)
            cel.Font.Bold = True
            cel.Font.Color = RGB(150, 75, 0)
        Case Else
            Dim ligneNum As Long: ligneNum = cel.Row
            If ligneNum Mod 2 = 0 Then
                cel.Interior.Color = RGB(235, 241, 255)
            Else
                cel.Interior.Color = RGB(255, 255, 255)
            End If
            cel.Font.Color = RGB(0, 0, 0)
            cel.Font.Bold = False
    End Select
End Sub

' ──────────────────────────────────────────────
' VALIDATION DE LA SAISIE
' ──────────────────────────────────────────────
Private Function ValiderSaisie() As Boolean
    ValiderSaisie = False

    ' Sélections obligatoires
    If m_projet = "" Then
        MsgBox "Veuillez sélectionner un projet.", vbExclamation: Exit Function
    End If
    If m_collab = "" Then
        MsgBox "Veuillez sélectionner un collaborateur.", vbExclamation: Exit Function
    End If
    If m_jourIdx = 0 Then
        MsgBox "Veuillez sélectionner un jour.", vbExclamation: Exit Function
    End If
    If m_ligneCollab = 0 Then
        MsgBox "Ligne collaborateur introuvable.", vbExclamation: Exit Function
    End If

    ' Si OFF ou CONGE → pas besoin d'horaires
    If chkOFF.Value Or chkCONGE.Value Then
        ValiderSaisie = True: Exit Function
    End If

    ' Sinon, entrée et sortie obligatoires
    If Not FormatHeureValide(Trim(txtEntree.Text)) Then
        MsgBox "Format entrée invalide. Utilisez HH:MM (ex: 08:00).", vbExclamation
        txtEntree.SetFocus: Exit Function
    End If
    If Not FormatHeureValide(Trim(txtSortie.Text)) Then
        MsgBox "Format sortie invalide. Utilisez HH:MM (ex: 17:00).", vbExclamation
        txtSortie.SetFocus: Exit Function
    End If

    ' Pause : si l'un est renseigné, les deux doivent l'être
    Dim pd As String: pd = Trim(txtPauseD.Text)
    Dim pf As String: pf = Trim(txtPauseF.Text)
    If (pd <> "" And pf = "") Or (pd = "" And pf <> "") Then
        MsgBox "Renseignez les deux bornes de pause ou laissez les deux vides.", vbExclamation
        Exit Function
    End If
    If pd <> "" Then
        If Not FormatHeureValide(pd) Or Not FormatHeureValide(pf) Then
            MsgBox "Format pause invalide. Utilisez HH:MM.", vbExclamation
            Exit Function
        End If
    End If

    ' Cohérence entrée < sortie
    If HeureEnMinutes_UF(Trim(txtEntree.Text)) >= HeureEnMinutes_UF(Trim(txtSortie.Text)) Then
        MsgBox "L'heure d'entrée doit être antérieure à l'heure de sortie.", vbExclamation
        Exit Function
    End If

    ValiderSaisie = True
End Function

' ──────────────────────────────────────────────
' HELPERS LOCAUX
' ──────────────────────────────────────────────
Private Function FormatHeureValide(h As String) As Boolean
    If h = "" Then FormatHeureValide = False: Exit Function
    If Not h Like "##:##" Then FormatHeureValide = False: Exit Function
    Dim p() As String: p = Split(h, ":")
    If Not IsNumeric(p(0)) Or Not IsNumeric(p(1)) Then FormatHeureValide = False: Exit Function
    If CInt(p(0)) > 23 Or CInt(p(1)) > 59 Then FormatHeureValide = False: Exit Function
    FormatHeureValide = True
End Function

Private Function HeureEnMinutes_UF(h As String) As Integer
    If h = "" Then HeureEnMinutes_UF = 0: Exit Function
    Dim p() As String: p = Split(h, ":")
    If UBound(p) < 1 Then HeureEnMinutes_UF = 0: Exit Function
    HeureEnMinutes_UF = CInt(p(0)) * 60 + CInt(p(1))
End Function

Private Function TrouverLigneCollab(ws As Worksheet, nomCollab As String) As Long
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 4 To lr
        If Trim(CStr(ws.Cells(i, 1).Value)) = nomCollab Then
            TrouverLigneCollab = i: Exit Function
        End If
    Next i
    TrouverLigneCollab = 0
End Function

Private Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not (ws Is Nothing)
End Function

Private Sub ViderFormulaire()
    txtEntree.Text = ""
    txtSortie.Text = ""
    txtPauseD.Text = ""
    txtPauseF.Text = ""
    chkOFF.Value = False
    chkCONGE.Value = False
    chkTT.Value = False
    chkRENFORT.Value = False
    chkSHIFT.Value = False
    txtEntree.Enabled = True
    txtSortie.Enabled = True
    txtPauseD.Enabled = True
    txtPauseF.Enabled = True
End Sub

' ──────────────────────────────────────────────
' BOUTON ANNULER
' ──────────────────────────────────────────────
Private Sub btnAnnuler_Click()
    Me.Hide
End Sub

' ──────────────────────────────────────────────
' BOUTON RESET — recharger les données initiales
' ──────────────────────────────────────────────
Private Sub btnReset_Click()
    If m_ligneCollab = 0 Or m_jourIdx = 0 Then Exit Sub
    ChargerDonneesCreneau
    lblInfo.Caption = "Données rechargées depuis la feuille."
End Sub

' ──────────────────────────────────────────────
' FERMETURE
' ──────────────────────────────────────────────
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub
