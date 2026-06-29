Attribute VB_Name = "ModuleUtils"
' ============================================================
' MODULE : ModuleUtils
' Fonctions utilitaires partagées par tous les modules
' ============================================================
Option Explicit

' ── Constantes globales ────────────────────────────────────
Public Const NOM_FEUILLE_PLANNING     As String = "PLANNING"
Public Const NOM_FEUILLE_ETAT         As String = "ETAT"
Public Const NOM_FEUILLE_UTILISATEURS As String = "Utilisateurs"

' Colonnes feuille PLANNING (format large, 1 ligne/collab/semaine)
Public Const COL_PL_SEMAINE      As Integer = 1
Public Const COL_PL_MATRICULE    As Integer = 2
Public Const COL_PL_NOM          As Integer = 3
Public Const COL_PL_PRENOM       As Integer = 4
Public Const COL_PL_NOMCOMPLET   As Integer = 5
Public Const COL_PL_DATEEMB      As Integer = 6
Public Const COL_PL_ACTIVITE     As Integer = 7
Public Const COL_PL_TEL          As Integer = 8
Public Const COL_PL_VILLE        As Integer = 9
Public Const COL_PL_POINTREPERE  As Integer = 10
Public Const COL_PL_ZONE         As Integer = 11
' Entrée/Sortie par jour : Lun=12/13 Mar=14/15 Mer=16/17 Jeu=18/19 Ven=20/21 Sam=22/23 Dim=24/25
Public Const COL_PL_LUN_E        As Integer = 12
Public Const COL_PL_DIM_S        As Integer = 25

' Heures affichées dans les tableaux
Public Const HEURE_MIN_ENTREE    As Integer = 7    ' 07H
Public Const HEURE_MAX_ENTREE    As Integer = 20   ' 20H
Public Const HEURE_MIN_SORTIE    As Integer = 7    ' 07H
Public Const HEURE_MAX_SORTIE    As Integer = 20   ' 20H

' Jours de la semaine (index 1=Lun … 7=Dim)
Public Const NB_JOURS            As Integer = 7

' ── Couleurs ───────────────────────────────────────────────
Public Const COULEUR_HEADER_BLEU  As Long = 2003199   ' RGB(31,73,125)
Public Const COULEUR_HEADER_CLAIR As Long = 12895428  ' RGB(68,114,196)
Public Const COULEUR_LIGNE_PAIRE  As Long = 16250858  ' RGB(235,241,255)
Public Const COULEUR_BLANC        As Long = 16777215
Public Const COULEUR_ENTREE       As Long = 13408767  ' RGB(197,225,165) vert entrée
Public Const COULEUR_SORTIE       As Long = 16757478  ' RGB(255,199,206) rose sortie
Public Const COULEUR_ZERO         As Long = 15921906  ' RGB(242,242,242) gris zéro

' ============================================================
' Retourne True si la feuille existe
' ============================================================
Public Function FeuilleExiste(nom As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nom)
    On Error GoTo 0
    FeuilleExiste = Not (ws Is Nothing)
End Function

' ============================================================
' Retourne la feuille, la crée si elle n'existe pas
' ============================================================
Public Function ObtenirOuCreerFeuille(nom As String) As Worksheet
    If FeuilleExiste(nom) Then
        Set ObtenirOuCreerFeuille = ThisWorkbook.Sheets(nom)
    Else
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = nom
        Set ObtenirOuCreerFeuille = ws
    End If
End Function

' ============================================================
' Extrait l'heure entière depuis une chaîne "HH:MM" ou "TT HH:MM"
' Retourne -1 si invalide ou OFF/CONGE
' ============================================================
Public Function ExtraireHeure(valeur As String) As Integer
    valeur = Trim(valeur)
    If valeur = "" Or valeur = "OFF" Or valeur = "CONGE" Then
        ExtraireHeure = -1: Exit Function
    End If
    ' Enlever préfixe TT
    If Left(valeur, 3) = "TT " Then valeur = Trim(Mid(valeur, 4))
    ' Extraire HH:MM
    Dim pos As Integer: pos = InStr(valeur, ":")
    If pos < 2 Then ExtraireHeure = -1: Exit Function
    Dim h As String: h = Left(valeur, pos - 1)
    h = Trim(h)
    If IsNumeric(h) Then
        ExtraireHeure = CInt(h)
    Else
        ExtraireHeure = -1
    End If
End Function

' ============================================================
' Retourne le nom du jour (1=Lundi … 7=Dimanche)
' ============================================================
Public Function NomJour(idx As Integer) As String
    Select Case idx
        Case 1: NomJour = "Lundi"
        Case 2: NomJour = "Mardi"
        Case 3: NomJour = "Mercredi"
        Case 4: NomJour = "Jeudi"
        Case 5: NomJour = "Vendredi"
        Case 6: NomJour = "Samedi"
        Case 7: NomJour = "Dimanche"
        Case Else: NomJour = ""
    End Select
End Function

' ============================================================
' Retourne l'indice jour (1=Lun) depuis le nom
' ============================================================
Public Function IndexJour(nomJour As String) As Integer
    Select Case UCase(Trim(nomJour))
        Case "LUNDI":    IndexJour = 1
        Case "MARDI":    IndexJour = 2
        Case "MERCREDI": IndexJour = 3
        Case "JEUDI":    IndexJour = 4
        Case "VENDREDI": IndexJour = 5
        Case "SAMEDI":   IndexJour = 6
        Case "DIMANCHE": IndexJour = 7
        Case Else:       IndexJour = 0
    End Select
End Function

' ============================================================
' Formate une heure entière en "07H", "08H" …
' ============================================================
Public Function FormatHeure(h As Integer) As String
    FormatHeure = Format(h, "00") & "H"
End Function

' ============================================================
' Applique la mise en forme d'un en-tête de section
' ============================================================
Public Sub StyleEnTete(cel As Range, Optional taille As Integer = 11)
    With cel
        .Font.Bold = True
        .Font.Size = taille
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = COULEUR_HEADER_BLEU
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' ============================================================
' Applique des bordures fines sur une plage
' ============================================================
Public Sub AppliquerBordures(rng As Range, Optional couleurBord As Long = -1)
    If couleurBord = -1 Then couleurBord = RGB(180, 180, 180)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = couleurBord
    End With
    With rng.Borders(xlEdgeLeft):   .Weight = xlMedium: .Color = COULEUR_HEADER_CLAIR: End With
    With rng.Borders(xlEdgeRight):  .Weight = xlMedium: .Color = COULEUR_HEADER_CLAIR: End With
    With rng.Borders(xlEdgeTop):    .Weight = xlMedium: .Color = COULEUR_HEADER_CLAIR: End With
    With rng.Borders(xlEdgeBottom): .Weight = xlMedium: .Color = COULEUR_HEADER_CLAIR: End With
End Sub

' ============================================================
' Désactive les mises à jour écran / calcul automatique
' ============================================================
Public Sub DebutCalcul()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
End Sub

' ============================================================
' Réactive les mises à jour écran / calcul automatique
' ============================================================
Public Sub FinCalcul()
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
