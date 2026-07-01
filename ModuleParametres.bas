Attribute VB_Name = "ModuleParametres"
Option Explicit

' ============================================================
' MODULE PARAMETRES - Generateur de planning
' Gere la feuille PARAMETRES (Cle / Valeur / Categorie / Libelle)
' et fournit GetParam / SetParam utilises par BOOM et UFParametres
' ============================================================

Public Const NOM_FEUILLE_PARAM As String = "PARAMETRES"

' ------------------------------------------------------------
' Catalogue complet des parametres : Categorie | Cle | Libelle | Valeur par defaut
' Source unique utilisee pour initialiser la feuille ET construire le UserForm
' ------------------------------------------------------------
Public Function CatalogueParametres() As Variant
    Dim c As Variant
    ReDim c(1 To 93, 1 To 4)
    c(1, 1) = "AFEDIM": c(1, 2) = "AFEDIM_LUNJEU_ENTREE": c(1, 3) = "Lun-Jeu Entree": c(1, 4) = "08:00"
    c(2, 1) = "AFEDIM": c(2, 2) = "AFEDIM_LUNJEU_SORTIE": c(2, 3) = "Lun-Jeu Sortie": c(2, 4) = "18:00"
    c(3, 1) = "AFEDIM": c(3, 2) = "AFEDIM_LUNJEU_PAUSED": c(3, 3) = "Lun-Jeu Pause debut": c(3, 4) = "13:00"
    c(4, 1) = "AFEDIM": c(4, 2) = "AFEDIM_LUNJEU_PAUSEF": c(4, 3) = "Lun-Jeu Pause fin": c(4, 4) = "14:00"
    c(5, 1) = "AFEDIM": c(5, 2) = "AFEDIM_VEN_ENTREE": c(5, 3) = "Vendredi Entree": c(5, 4) = "08:00"
    c(6, 1) = "AFEDIM": c(6, 2) = "AFEDIM_VEN_SORTIE": c(6, 3) = "Vendredi Sortie": c(6, 4) = "17:00"
    c(7, 1) = "AFEDIM": c(7, 2) = "AFEDIM_VEN_PAUSED": c(7, 3) = "Vendredi Pause debut": c(7, 4) = "13:00"
    c(8, 1) = "AFEDIM": c(8, 2) = "AFEDIM_VEN_PAUSEF": c(8, 3) = "Vendredi Pause fin": c(8, 4) = "14:00"
    c(9, 1) = "ACCESSIBILITE": c(9, 2) = "ACCESSIBILITE_LUNJEU_ENTREE": c(9, 3) = "Lun-Jeu Entree": c(9, 4) = "08:00"
    c(10, 1) = "ACCESSIBILITE": c(10, 2) = "ACCESSIBILITE_LUNJEU_SORTIE": c(10, 3) = "Lun-Jeu Sortie": c(10, 4) = "18:00"
    c(11, 1) = "ACCESSIBILITE": c(11, 2) = "ACCESSIBILITE_LUNJEU_PAUSED": c(11, 3) = "Lun-Jeu Pause debut": c(11, 4) = "13:00"
    c(12, 1) = "ACCESSIBILITE": c(12, 2) = "ACCESSIBILITE_LUNJEU_PAUSEF": c(12, 3) = "Lun-Jeu Pause fin": c(12, 4) = "14:00"
    c(13, 1) = "ACCESSIBILITE": c(13, 2) = "ACCESSIBILITE_VEN_ENTREE": c(13, 3) = "Vendredi Entree": c(13, 4) = "08:00"
    c(14, 1) = "ACCESSIBILITE": c(14, 2) = "ACCESSIBILITE_VEN_SORTIE": c(14, 3) = "Vendredi Sortie": c(14, 4) = "17:00"
    c(15, 1) = "ACCESSIBILITE": c(15, 2) = "ACCESSIBILITE_VEN_PAUSED": c(15, 3) = "Vendredi Pause debut": c(15, 4) = "13:00"
    c(16, 1) = "ACCESSIBILITE": c(16, 2) = "ACCESSIBILITE_VEN_PAUSEF": c(16, 3) = "Vendredi Pause fin": c(16, 4) = "14:00"
    c(17, 1) = "CM LEASING": c(17, 2) = "CMLEASING_LUNJEU_ENTREE": c(17, 3) = "Lun-Jeu Entree": c(17, 4) = "08:00"
    c(18, 1) = "CM LEASING": c(18, 2) = "CMLEASING_LUNJEU_SORTIE": c(18, 3) = "Lun-Jeu Sortie": c(18, 4) = "18:00"
    c(19, 1) = "CM LEASING": c(19, 2) = "CMLEASING_LUNJEU_PAUSED": c(19, 3) = "Lun-Jeu Pause debut": c(19, 4) = "13:00"
    c(20, 1) = "CM LEASING": c(20, 2) = "CMLEASING_LUNJEU_PAUSEF": c(20, 3) = "Lun-Jeu Pause fin": c(20, 4) = "14:00"
    c(21, 1) = "CM LEASING": c(21, 2) = "CMLEASING_VEN_ENTREE": c(21, 3) = "Vendredi Entree": c(21, 4) = "08:00"
    c(22, 1) = "CM LEASING": c(22, 2) = "CMLEASING_VEN_SORTIE": c(22, 3) = "Vendredi Sortie": c(22, 4) = "17:00"
    c(23, 1) = "CM LEASING": c(23, 2) = "CMLEASING_VEN_PAUSED": c(23, 3) = "Vendredi Pause debut": c(23, 4) = "13:00"
    c(24, 1) = "CM LEASING": c(24, 2) = "CMLEASING_VEN_PAUSEF": c(24, 3) = "Vendredi Pause fin": c(24, 4) = "14:00"
    c(25, 1) = "GLF": c(25, 2) = "GLF_LUNJEU_ENTREE": c(25, 3) = "Lun-Jeu Entree": c(25, 4) = "08:00"
    c(26, 1) = "GLF": c(26, 2) = "GLF_LUNJEU_SORTIE": c(26, 3) = "Lun-Jeu Sortie": c(26, 4) = "18:00"
    c(27, 1) = "GLF": c(27, 2) = "GLF_VEN_ENTREE": c(27, 3) = "Vendredi Entree": c(27, 4) = "08:00"
    c(28, 1) = "GLF": c(28, 2) = "GLF_VEN_SORTIE": c(28, 3) = "Vendredi Sortie": c(28, 4) = "17:00"
    c(29, 1) = "GLF": c(29, 2) = "GLF_VAGUES": c(29, 3) = "Vagues pause (liste separee par virgules)": c(29, 4) = "12:00,12:30,13:00,13:30,14:00"
    c(30, 1) = "GLF": c(30, 2) = "GLF_PAUSE_DUREE_MIN": c(30, 3) = "Duree pause (minutes)": c(30, 4) = "60"
    c(31, 1) = "EBRA": c(31, 2) = "EBRA_SEM_ENTREE": c(31, 3) = "Lun-Ven Entree": c(31, 4) = "07:00"
    c(32, 1) = "EBRA": c(32, 2) = "EBRA_SEM_SORTIE": c(32, 3) = "Lun-Ven Sortie": c(32, 4) = "16:00"
    c(33, 1) = "EBRA": c(33, 2) = "EBRA_SAM_ENTREE": c(33, 3) = "Samedi Entree": c(33, 4) = "07:00"
    c(34, 1) = "EBRA": c(34, 2) = "EBRA_SAM_SORTIE": c(34, 3) = "Samedi Sortie": c(34, 4) = "11:00"
    c(35, 1) = "EBRA": c(35, 2) = "EBRA_VAGUES": c(35, 3) = "Vagues pause (liste separee par virgules)": c(35, 4) = "11:00,11:30,12:00,12:30,13:00"
    c(36, 1) = "EBRA": c(36, 2) = "EBRA_PAUSE_DUREE_MIN": c(36, 3) = "Duree pause (minutes)": c(36, 4) = "60"
    c(37, 1) = "GOOGLE LEADS": c(37, 2) = "GL_SHIFT1_ENTREE": c(37, 3) = "Shift 1 Entree": c(37, 4) = "07:00"
    c(38, 1) = "GOOGLE LEADS": c(38, 2) = "GL_SHIFT1_SORTIE": c(38, 3) = "Shift 1 Sortie normale": c(38, 4) = "17:00"
    c(39, 1) = "GOOGLE LEADS": c(39, 2) = "GL_SHIFT1_SORTIE_REDUITE": c(39, 3) = "Shift 1 Sortie reduite (-2h)": c(39, 4) = "16:00"
    c(40, 1) = "GOOGLE LEADS": c(40, 2) = "GL_SHIFT2_ENTREE": c(40, 3) = "Shift 2 Entree": c(40, 4) = "08:00"
    c(41, 1) = "GOOGLE LEADS": c(41, 2) = "GL_SHIFT2_SORTIE": c(41, 3) = "Shift 2 Sortie normale": c(41, 4) = "18:00"
    c(42, 1) = "GOOGLE LEADS": c(42, 2) = "GL_SHIFT2_SORTIE_REDUITE": c(42, 3) = "Shift 2 Sortie reduite (-2h)": c(42, 4) = "17:00"
    c(43, 1) = "GOOGLE LEADS": c(43, 2) = "GL_SHIFT3_ENTREE": c(43, 3) = "Shift 3 Entree": c(43, 4) = "09:00"
    c(44, 1) = "GOOGLE LEADS": c(44, 2) = "GL_SHIFT3_SORTIE": c(44, 3) = "Shift 3 Sortie normale": c(44, 4) = "19:00"
    c(45, 1) = "GOOGLE LEADS": c(45, 2) = "GL_SHIFT3_SORTIE_REDUITE": c(45, 3) = "Shift 3 Sortie reduite (-2h)": c(45, 4) = "18:00"
    c(46, 1) = "GOOGLE LEADS": c(46, 2) = "GL_SHIFT4_ENTREE": c(46, 3) = "Shift 4 Entree": c(46, 4) = "10:00"
    c(47, 1) = "GOOGLE LEADS": c(47, 2) = "GL_SHIFT4_SORTIE": c(47, 3) = "Shift 4 Sortie normale": c(47, 4) = "20:00"
    c(48, 1) = "GOOGLE LEADS": c(48, 2) = "GL_SHIFT4_SORTIE_REDUITE": c(48, 3) = "Shift 4 Sortie reduite (-2h)": c(48, 4) = "19:00"
    c(49, 1) = "GOOGLE LEADS": c(49, 2) = "GL_SHIFT5_ENTREE": c(49, 3) = "Shift 5 Entree": c(49, 4) = "07:00"
    c(50, 1) = "GOOGLE LEADS": c(50, 2) = "GL_SHIFT5_SORTIE": c(50, 3) = "Shift 5 Sortie normale": c(50, 4) = "17:00"
    c(51, 1) = "GOOGLE LEADS": c(51, 2) = "GL_SHIFT5_SORTIE_REDUITE": c(51, 3) = "Shift 5 Sortie reduite (-2h)": c(51, 4) = "16:00"
    c(52, 1) = "GOOGLE LEADS": c(52, 2) = "GL_PAUSE_OFFSET_MIN": c(52, 3) = "Pause : decalage apres entree (min)": c(52, 4) = "300"
    c(53, 1) = "GOOGLE LEADS": c(53, 2) = "GL_PAUSE_DUREE_MIN": c(53, 3) = "Duree pause (minutes)": c(53, 4) = "60"
    c(54, 1) = "GOOGLE LEADS": c(54, 2) = "GL_QUOTA_OFF_SEMAINE": c(54, 3) = "Quota OFF/jour Lun-Ven": c(54, 4) = "5"
    c(55, 1) = "GOOGLE LEADS": c(55, 2) = "GL_QUOTA_OFF_SAMEDI": c(55, 3) = "Quota OFF Samedi": c(55, 4) = "6"
    c(56, 1) = "GOOGLE LEADS": c(56, 2) = "GL_QUOTA_OFF_DIMANCHE": c(56, 3) = "Quota OFF Dimanche": c(56, 4) = "7"
    c(57, 1) = "GOOGLE LEADS": c(57, 2) = "GL_REDUIT_DIVISEUR": c(57, 3) = "Diviseur limite shifts reduits/jour": c(57, 4) = "7"
    c(58, 1) = "TLV": c(58, 2) = "TLV_NOM1": c(58, 3) = "Nom agent groupe fixe 1": c(58, 4) = "SYLLA SOKHNA SAFIETOU"
    c(59, 1) = "TLV": c(59, 2) = "TLV_NOM2": c(59, 3) = "Nom agent groupe fixe 2": c(59, 4) = "DIOP MAMADOU MOUSTAPHA DOKY"
    c(60, 1) = "TLV": c(60, 2) = "TLV_NOM3": c(60, 3) = "Nom agent rotation 1": c(60, 4) = "ABDELAOUI KHADIJA"
    c(61, 1) = "TLV": c(61, 2) = "TLV_NOM4": c(61, 3) = "Nom agent rotation 2": c(61, 4) = "AZIANE YASSINE"
    c(62, 1) = "TLV": c(62, 2) = "TLV_G1_ENTREE": c(62, 3) = "Groupe fixe Lun-Ven Entree": c(62, 4) = "08:00"
    c(63, 1) = "TLV": c(63, 2) = "TLV_G1_SORTIE": c(63, 3) = "Groupe fixe Lun-Ven Sortie": c(63, 4) = "17:00"
    c(64, 1) = "TLV": c(64, 2) = "TLV_G1_PAUSED": c(64, 3) = "Groupe fixe Pause debut": c(64, 4) = "13:00"
    c(65, 1) = "TLV": c(65, 2) = "TLV_G1_PAUSEF": c(65, 3) = "Groupe fixe Pause fin": c(65, 4) = "14:00"
    c(66, 1) = "TLV": c(66, 2) = "TLV_G2_LUN_ENTREE": c(66, 3) = "Groupe rotation Lundi Entree": c(66, 4) = "08:00"
    c(67, 1) = "TLV": c(67, 2) = "TLV_G2_LUN_SORTIE": c(67, 3) = "Groupe rotation Lundi Sortie": c(67, 4) = "17:00"
    c(68, 1) = "TLV": c(68, 2) = "TLV_G2_MARMER_ENTREE": c(68, 3) = "Groupe rotation Mar-Mer Entree": c(68, 4) = "08:00"
    c(69, 1) = "TLV": c(69, 2) = "TLV_G2_MARMER_SORTIE": c(69, 3) = "Groupe rotation Mar-Mer Sortie": c(69, 4) = "18:00"
    c(70, 1) = "TLV": c(70, 2) = "TLV_G2_JEUVEN_ENTREE": c(70, 3) = "Groupe rotation Jeu-Ven Entree": c(70, 4) = "08:00"
    c(71, 1) = "TLV": c(71, 2) = "TLV_G2_JEUVEN_SORTIE": c(71, 3) = "Groupe rotation Jeu-Ven Sortie": c(71, 4) = "17:00"
    c(72, 1) = "TLV": c(72, 2) = "TLV_G2_SAM_ENTREE": c(72, 3) = "Groupe rotation Samedi Entree": c(72, 4) = "08:00"
    c(73, 1) = "TLV": c(73, 2) = "TLV_G2_SAM_SORTIE": c(73, 3) = "Groupe rotation Samedi Sortie": c(73, 4) = "14:00"
    c(74, 1) = "TLV": c(74, 2) = "TLV_G2_PAUSED": c(74, 3) = "Groupe rotation Pause debut": c(74, 4) = "13:00"
    c(75, 1) = "TLV": c(75, 2) = "TLV_G2_PAUSEF": c(75, 3) = "Groupe rotation Pause fin": c(75, 4) = "14:00"
    c(76, 1) = "TLV": c(76, 2) = "TLV_FALLBACK_ENTREE": c(76, 3) = "Autre agent TLV Entree": c(76, 4) = "08:00"
    c(77, 1) = "TLV": c(77, 2) = "TLV_FALLBACK_SORTIE": c(77, 3) = "Autre agent TLV Sortie": c(77, 4) = "17:00"
    c(78, 1) = "TLV": c(78, 2) = "TLV_FALLBACK_PAUSED": c(78, 3) = "Autre agent TLV Pause debut": c(78, 4) = "13:00"
    c(79, 1) = "TLV": c(79, 2) = "TLV_FALLBACK_PAUSEF": c(79, 3) = "Autre agent TLV Pause fin": c(79, 4) = "14:00"
    c(80, 1) = "FACTO": c(80, 2) = "FACTO_SHIFT1_ENTREE": c(80, 3) = "Shift 1 Entree": c(80, 4) = "07:00"
    c(81, 1) = "FACTO": c(81, 2) = "FACTO_SHIFT1_SORTIE": c(81, 3) = "Shift 1 Sortie": c(81, 4) = "17:00"
    c(82, 1) = "FACTO": c(82, 2) = "FACTO_SHIFT2_ENTREE": c(82, 3) = "Shift 2 Entree": c(82, 4) = "08:00"
    c(83, 1) = "FACTO": c(83, 2) = "FACTO_SHIFT2_SORTIE": c(83, 3) = "Shift 2 Sortie": c(83, 4) = "18:00"
    c(84, 1) = "FACTO": c(84, 2) = "FACTO_VEN_REDUCTION_MIN": c(84, 3) = "Reduction Vendredi (min)": c(84, 4) = "60"
    c(85, 1) = "FACTO": c(85, 2) = "FACTO_PAUSE_OFFSET_MIN": c(85, 3) = "Pause : decalage apres entree (min)": c(85, 4) = "300"
    c(86, 1) = "FACTO": c(86, 2) = "FACTO_PAUSE_DUREE_MIN": c(86, 3) = "Duree pause (minutes)": c(86, 4) = "60"
    c(87, 1) = "DAC": c(87, 2) = "DAC_SHIFT1_ENTREE": c(87, 3) = "Shift 1 Entree": c(87, 4) = "07:00"
    c(88, 1) = "DAC": c(88, 2) = "DAC_SHIFT1_SORTIE": c(88, 3) = "Shift 1 Sortie": c(88, 4) = "17:00"
    c(89, 1) = "DAC": c(89, 2) = "DAC_SHIFT2_ENTREE": c(89, 3) = "Shift 2 Entree": c(89, 4) = "08:00"
    c(90, 1) = "DAC": c(90, 2) = "DAC_SHIFT2_SORTIE": c(90, 3) = "Shift 2 Sortie": c(90, 4) = "18:00"
    c(91, 1) = "DAC": c(91, 2) = "DAC_VEN_REDUCTION_MIN": c(91, 3) = "Reduction Vendredi (min)": c(91, 4) = "60"
    c(92, 1) = "DAC": c(92, 2) = "DAC_PAUSE_OFFSET_MIN": c(92, 3) = "Pause : decalage apres entree (min)": c(92, 4) = "300"
    c(93, 1) = "DAC": c(93, 2) = "DAC_PAUSE_DUREE_MIN": c(93, 3) = "Duree pause (minutes)": c(93, 4) = "60"
    CatalogueParametres = c
End Function

' ============================================================
' INITIALISATION DE LA FEUILLE PARAMETRES
' Cree la feuille si absente, ajoute les cles manquantes avec
' leur valeur par defaut (ne touche jamais aux valeurs existantes)
' ============================================================
Sub InitialiserFeuilleParametres()
    Dim ws As Worksheet
    If Not FeuilleExiste(NOM_FEUILLE_PARAM) Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = NOM_FEUILLE_PARAM
    Else
        Set ws = ThisWorkbook.Sheets(NOM_FEUILLE_PARAM)
    End If

    If ws.Cells(1, 1).Value <> "Cle" Then
        ws.Cells(1, 1).Value = "Cle"
        ws.Cells(1, 2).Value = "Valeur"
        ws.Cells(1, 3).Value = "Categorie"
        ws.Cells(1, 4).Value = "Libelle"
        With ws.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(31, 73, 125)
            .Font.Color = RGB(255, 255, 255)
        End With
        ws.Columns("A").ColumnWidth = 32
        ws.Columns("B").ColumnWidth = 40
        ws.Columns("C").ColumnWidth = 20
        ws.Columns("D").ColumnWidth = 45
    End If

    Dim cat As Variant: cat = CatalogueParametres()
    Dim i As Long
    For i = 1 To UBound(cat, 1)
        Dim lr As Long: lr = TrouverLigneParam(ws, CStr(cat(i, 2)))
        If lr = 0 Then
            lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
            ws.Cells(lr, 1).Value = cat(i, 2)
            ws.Cells(lr, 2).Value = cat(i, 4)
            ws.Cells(lr, 3).Value = cat(i, 1)
            ws.Cells(lr, 4).Value = cat(i, 3)
        Else
            ' Categorie/Libelle peuvent avoir change de puis une version anterieure : on les rafraichit
            ws.Cells(lr, 3).Value = cat(i, 1)
            ws.Cells(lr, 4).Value = cat(i, 3)
        End If
    Next i
End Sub

Function TrouverLigneParam(ws As Worksheet, cle As String) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = cle Then TrouverLigneParam = i: Exit Function
    Next i
    TrouverLigneParam = 0
End Function

' ============================================================
' LECTURE / ECRITURE D'UN PARAMETRE
' ============================================================
Function GetParam(cle As String, Optional defaut As String = "") As String
    Static ws As Worksheet
    If ws Is Nothing Then
        If Not FeuilleExiste(NOM_FEUILLE_PARAM) Then InitialiserFeuilleParametres
        Set ws = ThisWorkbook.Sheets(NOM_FEUILLE_PARAM)
    End If
    Dim lr As Long: lr = TrouverLigneParam(ws, cle)
    If lr = 0 Then
        GetParam = defaut
    Else
        Dim v As String: v = CStr(ws.Cells(lr, 2).Value)
        GetParam = IIf(v = "", defaut, v)
    End If
End Function

' Variante numerique pratique (quotas, minutes, diviseurs...)
Function GetParamNum(cle As String, Optional defaut As Double = 0) As Double
    Dim v As String: v = GetParam(cle, CStr(defaut))
    If IsNumeric(v) Then GetParamNum = CDbl(v) Else GetParamNum = defaut
End Function

Sub SetParam(cle As String, valeur As String)
    If Not FeuilleExiste(NOM_FEUILLE_PARAM) Then InitialiserFeuilleParametres
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(NOM_FEUILLE_PARAM)
    Dim lr As Long: lr = TrouverLigneParam(ws, cle)
    If lr = 0 Then
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(lr, 1).Value = cle
    End If
    ws.Cells(lr, 2).Value = valeur
End Sub

' Renvoie GLF_VAGUES / EBRA_VAGUES ("12:00,12:30,...") sous forme de tableau 1..5
Function GetParamListe(cle As String, defaut As String) As Variant
    Dim v As String: v = GetParam(cle, defaut)
    GetParamListe = Split(v, ",")
End Function

' ============================================================
' REINITIALISATION AUX VALEURS PAR DEFAUT
' ============================================================
Sub ReinitialiserParametresDefaut()
    If Not FeuilleExiste(NOM_FEUILLE_PARAM) Then InitialiserFeuilleParametres
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(NOM_FEUILLE_PARAM)
    Dim cat As Variant: cat = CatalogueParametres()
    Dim i As Long
    For i = 1 To UBound(cat, 1)
        SetParam CStr(cat(i, 2)), CStr(cat(i, 4))
    Next i
End Sub

Sub OuvrirUFParametres()
    If Not FeuilleExiste(NOM_FEUILLE_PARAM) Then InitialiserFeuilleParametres
    UFParametres.Show
End Sub

