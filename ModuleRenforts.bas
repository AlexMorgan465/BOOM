Attribute VB_Name = "ModuleRenforts"
' ============================================================
' MODULE : ModuleRenforts  [RÉSERVÉ — ARCHITECTURE FUTURE]
' Calcul des besoins de renfort, sous-effectifs et taux d'occupation
'
' Ce module exploitera les données de ModulePlanning et
' ModuleStatistiques pour calculer :
'   - Sous-effectifs par créneau horaire
'   - Capacité maximale par ville/jour
'   - Taux d'occupation (effectif présent / capacité max)
'   - Proposition automatique de renforts
'   - Suivi des renforts passés (équité)
' ============================================================
Option Explicit

' ── Seuils de sous-effectif (à paramétrer) ─────────────────
' Exemple : si effectif présent < SeuilMinimum pour un créneau
'           → le créneau est marqué "sous-effectif"
Private Const SEUIL_MINIMUM_DEFAUT As Integer = 5

' ── Structure Créneau ──────────────────────────────────────
Public Type Creneau
    Jour        As String
    Ville       As String
    Heure       As Integer
    Effectif    As Integer
    Capacite    As Integer
    SousEffectif As Boolean
    TauxOccup   As Double
End Type

' ============================================================
' [FUTUR] Calculer les créneaux sous-effectif
' ============================================================
Public Sub CalculerSousEffectifs()
    ' TODO : implémenter quand la capacité max par ville/créneau
    '        sera définie dans une feuille de paramétrage
    MsgBox "Module Renforts : fonctionnalité à venir.", vbInformation
End Sub

' ============================================================
' [FUTUR] Proposer des renforts automatiques
' ============================================================
Public Sub ProposerRenforts()
    ' TODO : utiliser la feuille BESOINS existante et les
    '        critères d'éligibilité (RenforcPress, RenforcItaly)
    MsgBox "Module Renforts : fonctionnalité à venir.", vbInformation
End Sub

' ============================================================
' [FUTUR] Calculer taux d'occupation
' ============================================================
Public Function TauxOccupation(jour As String, ville As String, heure As Integer, capacite As Integer) As Double
    If capacite = 0 Then TauxOccupation = 0: Exit Function
    Dim effectif As Integer
    effectif = ModuleStatistiques.EffectifPresent(jour, ville, heure)
    TauxOccupation = Round(effectif / capacite * 100, 1)
End Function
