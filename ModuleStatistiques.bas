Attribute VB_Name = "ModuleStatistiques"
' ============================================================
' MODULE : ModuleStatistiques
' Calcul des statistiques d'entrées et sorties
' par Jour + Ville + Heure, via Scripting.Dictionary
' ============================================================
Option Explicit

' ── Structure de résultat ──────────────────────────────────
' Clé dictionnaire : "JOUR|VILLE|HEURE"  (ex: "Lundi|Rabat|8")
' Valeur : Integer (nombre d'agents)

Private m_dictEntrees  As Object   ' Scripting.Dictionary
Private m_dictSorties  As Object
Private m_estCalcule   As Boolean

' ============================================================
' CALCUL PRINCIPAL
' Remplit les deux dictionnaires depuis les données PLANNING
' ============================================================
Public Sub CalculerStatistiques()
    m_estCalcule = False

    ' Forcer rechargement des données
    ModulePlanning.Invalider
    ModulePlanning.ChargerDonnees

    If Not ModulePlanning.EstCharge() Then Exit Sub

    ' Initialiser les dictionnaires
    Set m_dictEntrees = CreateObject("Scripting.Dictionary")
    Set m_dictSorties = CreateObject("Scripting.Dictionary")
    m_dictEntrees.CompareMode = vbTextCompare
    m_dictSorties.CompareMode = vbTextCompare

    Dim donnees() As EnregistrementJour
    donnees = ModulePlanning.GetDonnees()
    Dim nb As Long: nb = ModulePlanning.GetNbDonnees()

    Dim i As Long
    For i = 1 To nb
        With donnees(i)
            ' Ignorer les OFF, CONGE et lignes sans ville
            If Not .EstOFF And Not .EstConge And .Ville <> "" Then

                ' ── Entrées ──
                If .HeureEntree >= HEURE_MIN_ENTREE And .HeureEntree <= HEURE_MAX_ENTREE Then
                    Dim cleE As String
                    cleE = .NomJourStr & "|" & .Ville & "|" & CStr(.HeureEntree)
                    If m_dictEntrees.Exists(cleE) Then
                        m_dictEntrees(cleE) = m_dictEntrees(cleE) + 1
                    Else
                        m_dictEntrees.Add cleE, 1
                    End If
                End If

                ' ── Sorties ──
                If .HeureSortie >= HEURE_MIN_SORTIE And .HeureSortie <= HEURE_MAX_SORTIE Then
                    Dim cleS As String
                    cleS = .NomJourStr & "|" & .Ville & "|" & CStr(.HeureSortie)
                    If m_dictSorties.Exists(cleS) Then
                        m_dictSorties(cleS) = m_dictSorties(cleS) + 1
                    Else
                        m_dictSorties.Add cleS, 1
                    End If
                End If

            End If
        End With
    Next i

    m_estCalcule = True
End Sub

' ============================================================
' GETTERS
' ============================================================
Public Function GetEntrees() As Object
    Set GetEntrees = m_dictEntrees
End Function

Public Function GetSorties() As Object
    Set GetSorties = m_dictSorties
End Function

Public Function GetValeurEntree(nomJour As String, ville As String, heure As Integer) As Integer
    Dim cle As String: cle = nomJour & "|" & ville & "|" & CStr(heure)
    If m_dictEntrees.Exists(cle) Then
        GetValeurEntree = m_dictEntrees(cle)
    Else
        GetValeurEntree = 0
    End If
End Function

Public Function GetValeurSortie(nomJour As String, ville As String, heure As Integer) As Integer
    Dim cle As String: cle = nomJour & "|" & ville & "|" & CStr(heure)
    If m_dictSorties.Exists(cle) Then
        GetValeurSortie = m_dictSorties(cle)
    Else
        GetValeurSortie = 0
    End If
End Function

Public Function EstCalcule() As Boolean
    EstCalcule = m_estCalcule
End Function

' ============================================================
' FUTUR : Effectifs présents heure par heure
' Retourne le nombre d'agents présents entre HeureEntree et HeureSortie
' à une heure H donnée, pour un jour et une ville
' (architecture préparée, non encore exposée dans les tableaux)
' ============================================================
Public Function EffectifPresent(nomJour As String, ville As String, heure As Integer) As Integer
    If Not m_estCalcule Then EffectifPresent = 0: Exit Function
    Dim donnees() As EnregistrementJour
    donnees = ModulePlanning.GetDonnees()
    Dim nb As Long: nb = ModulePlanning.GetNbDonnees()
    Dim count As Integer: count = 0
    Dim i As Long
    For i = 1 To nb
        With donnees(i)
            If Not .EstOFF And Not .EstConge _
               And .NomJourStr = nomJour And .Ville = ville _
               And .HeureEntree >= 0 And .HeureSortie > 0 _
               And heure >= .HeureEntree And heure < .HeureSortie Then
                count = count + 1
            End If
        End With
    Next i
    EffectifPresent = count
End Function
