Attribute VB_Name = "ModulePlanning"
' ============================================================
' MODULE : ModulePlanning
' Lecture et parsing de la feuille PLANNING en mémoire
' Structure PLANNING (25 colonnes) :
'   1=Semaine | 2=Matricule | 3=NOM | 4=PRENOM | 5=NOM COMPLET
'   6=Date embauche | 7=Activité | 8=Téléphone | 9=Ville
'   10=Point de repère | 11=Zone
'   12/13=LUN E/S | 14/15=MAR E/S | 16/17=MER E/S | 18/19=JEU E/S
'   20/21=VEN E/S | 22/23=SAM E/S | 24/25=DIM E/S
' ============================================================
Option Explicit

' ── Structure d'un enregistrement agent/jour ───────────────
Public Type EnregistrementJour
    NomComplet  As String
    Matricule   As String
    Ville       As String
    Zone        As String
    PointRepere As String
    Activite    As String
    Semaine     As Integer
    Jour        As Integer      ' 1=Lun … 7=Dim
    NomJourStr  As String
    HeureEntree As Integer      ' -1 si OFF/CONGE
    HeureSortie As Integer      ' -1 si OFF/CONGE
    EstOFF      As Boolean
    EstConge    As Boolean
    EstTT       As Boolean
    CleJourVille As String      ' ex: "Lundi|Rabat"
End Type

' ── Cache en mémoire ───────────────────────────────────────
Private m_donnees()     As EnregistrementJour
Private m_nbDonnees     As Long
Private m_villes()      As String
Private m_nbVilles      As Integer
Private m_estCharge     As Boolean

' ============================================================
' POINT D'ENTRÉE : charger toutes les données PLANNING
' Lecture en tableau VBA (très rapide même sur milliers de lignes)
' ============================================================
Public Sub ChargerDonnees()
    m_estCharge = False
    m_nbDonnees = 0
    m_nbVilles = 0

    If Not FeuilleExiste(NOM_FEUILLE_PLANNING) Then
        MsgBox "Feuille PLANNING introuvable.", vbCritical: Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(NOM_FEUILLE_PLANNING)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_PL_NOMCOMPLET).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Lecture en masse dans un tableau 2D (1 seul accès disque)
    Dim data As Variant
    data = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, COL_PL_DIM_S)).Value

    Dim nbLignes As Long: nbLignes = UBound(data, 1)

    ' Pré-allouer : 7 jours * nbLignes au maximum
    ReDim m_donnees(1 To nbLignes * 7)

    ' Dictionnaire pour dédupliquer les villes
    Dim dictVilles As Object
    Set dictVilles = CreateObject("Scripting.Dictionary")
    dictVilles.CompareMode = vbTextCompare

    Dim i As Long, j As Integer
    For i = 1 To nbLignes
        Dim nomComplet  As String: nomComplet  = Trim(CStr(data(i, COL_PL_NOMCOMPLET)))
        If nomComplet = "" Then GoTo NextLigne

        Dim matricule   As String: matricule   = Trim(CStr(data(i, COL_PL_MATRICULE)))
        Dim ville       As String: ville       = Trim(CStr(data(i, COL_PL_VILLE)))
        Dim zone        As String: zone        = Trim(CStr(data(i, COL_PL_ZONE)))
        Dim pointRepere As String: pointRepere = Trim(CStr(data(i, COL_PL_POINTREPERE)))
        Dim activite    As String: activite    = Trim(CStr(data(i, COL_PL_ACTIVITE)))
        Dim semaine     As Integer
        If IsNumeric(data(i, COL_PL_SEMAINE)) Then semaine = CInt(data(i, COL_PL_SEMAINE))

        ' Collecter les villes uniques
        If ville <> "" And Not dictVilles.Exists(ville) Then
            dictVilles.Add ville, True
        End If

        ' Traiter chaque jour (colonnes paires=Entrée, impaires=Sortie)
        ' Lun=col12/13, Mar=14/15 ... Dim=24/25
        For j = 1 To 7
            Dim colE As Integer: colE = 10 + (j * 2)   ' 12,14,16,18,20,22,24
            Dim colS As Integer: colS = colE + 1

            Dim valE As String: valE = Trim(CStr(data(i, colE)))
            Dim valS As String: valS = Trim(CStr(data(i, colS)))

            m_nbDonnees = m_nbDonnees + 1
            With m_donnees(m_nbDonnees)
                .NomComplet   = nomComplet
                .Matricule    = matricule
                .Ville        = ville
                .Zone         = zone
                .PointRepere  = pointRepere
                .Activite     = activite
                .Semaine      = semaine
                .Jour         = j
                .NomJourStr   = NomJour(j)
                .EstOFF        = (valE = "OFF" Or valE = "")
                .EstConge      = (UCase(valE) = "CONGE")
                .EstTT         = (Left(valE, 3) = "TT ")
                .CleJourVille  = NomJour(j) & "|" & ville

                If Not .EstOFF And Not .EstConge Then
                    .HeureEntree = ExtraireHeure(valE)
                    .HeureSortie = ExtraireHeure(valS)
                Else
                    .HeureEntree = -1
                    .HeureSortie = -1
                End If
            End With
        Next j
NextLigne:
    Next i

    ' Construire tableau des villes triées
    Dim clefVilles As Variant: clefVilles = dictVilles.Keys
    m_nbVilles = dictVilles.Count
    ReDim m_villes(1 To IIf(m_nbVilles = 0, 1, m_nbVilles))
    Dim v As Integer
    For v = 0 To m_nbVilles - 1
        m_villes(v + 1) = clefVilles(v)
    Next v
    ' Tri alphabétique des villes
    TrierVilles

    m_estCharge = True
End Sub

' ============================================================
' Accès aux données chargées
' ============================================================
Public Function GetDonnees() As EnregistrementJour()
    If Not m_estCharge Then ChargerDonnees
    GetDonnees = m_donnees
End Function

Public Function GetNbDonnees() As Long
    GetNbDonnees = m_nbDonnees
End Function

Public Function GetVilles() As String()
    GetVilles = m_villes
End Function

Public Function GetNbVilles() As Integer
    GetNbVilles = m_nbVilles
End Function

Public Function EstCharge() As Boolean
    EstCharge = m_estCharge
End Function

Public Sub Invalider()
    m_estCharge = False
End Sub

' ============================================================
' Tri à bulles des villes (ordre alphabétique)
' ============================================================
Private Sub TrierVilles()
    Dim a As Integer, b As Integer, tmp As String
    For a = 1 To m_nbVilles - 1
        For b = a + 1 To m_nbVilles
            If m_villes(a) > m_villes(b) Then
                tmp = m_villes(a): m_villes(a) = m_villes(b): m_villes(b) = tmp
            End If
        Next b
    Next a
End Sub
