VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFUtilisateurs 
   Caption         =   "UserForm1"
   ClientHeight    =   7344
   ClientLeft      =   -12
   ClientTop       =   -24
   ClientWidth     =   22464
   OleObjectBlob   =   "UFUtilisateurs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFUtilisateurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_ligneSelectionnee As Long
Private m_modeAjout As Boolean

' ============================================================
Private Sub UserForm_Initialize()
    ' Remplir la liste des projets
    cboProjet.Clear
    Dim projets As Variant
    projets = Array("AFEDIM", "ACCESSIBILITE", "CM Leasing", "GLF", "EBRA", _
                    "EBRA PRESSE", "GOOGLE LEADS", "TLV", "TELEVENTE", "FACTO", "DAC")
    Dim p As Variant
    For Each p In projets
        cboProjet.AddItem p
    Next p

    ' Transport
    cboTransport.Clear
    cboTransport.AddItem "OUI"
    cboTransport.AddItem "NON"
    cboTransport.Text = "NON"

    m_ligneSelectionnee = 0
    m_modeAjout = False
    ChargerListe
    VerrouillerFormulaire True
End Sub

' --- Charger la liste des collaborateurs ---
Private Sub ChargerListe()
    lstCollabs.Clear
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) <> "" Then
            lstCollabs.AddItem ws.Cells(i, 1).Value & "  [" & ws.Cells(i, 2).Value & "]"
            lstCollabs.List(lstCollabs.ListCount - 1, 0) = ws.Cells(i, 1).Value & "  [" & ws.Cells(i, 2).Value & "]"
        End If
    Next i
End Sub

' --- Sélection dans la liste ---
Private Sub lstCollabs_Click()
    If lstCollabs.ListIndex < 0 Then Exit Sub
    m_modeAjout = False
    m_ligneSelectionnee = lstCollabs.ListIndex + 2  ' +2 car ligne 1 = entête

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")

    txtNom.Text = CStr(ws.Cells(m_ligneSelectionnee, 1).Value)
    cboProjet.Text = CStr(ws.Cells(m_ligneSelectionnee, 2).Value)
    txtVille.Text = CStr(ws.Cells(m_ligneSelectionnee, 3).Value)
    txtZone.Text = CStr(ws.Cells(m_ligneSelectionnee, 4).Value)

    ' Congé (col 5-7)
    chkConge.Value = (UCase(Trim(ws.Cells(m_ligneSelectionnee, 5).Value)) = "OUI")
    txtCongeD.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 6).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 6).Value, "dd/mm/yyyy"), "")
    txtCongeF.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 7).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 7).Value, "dd/mm/yyyy"), "")

    ' Transport (col 8)
    cboTransport.Text = CStr(ws.Cells(m_ligneSelectionnee, 8).Value)

    ' TT (col 9-11)
    chkTT.Value = (UCase(Trim(ws.Cells(m_ligneSelectionnee, 9).Value)) = "OUI")
    txtTTD.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 10).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 10).Value, "dd/mm/yyyy"), "")
    txtTTF.Text = IIf(IsDate(ws.Cells(m_ligneSelectionnee, 11).Value), _
                           Format(ws.Cells(m_ligneSelectionnee, 11).Value, "dd/mm/yyyy"), "")

    ' Renforts (col 12-13)
    chkRenfortP.Value = (UCase(Trim(ws.Cells(m_ligneSelectionnee, 12).Value)) = "OUI")
    chkRenfortI.Value = (UCase(Trim(ws.Cells(m_ligneSelectionnee, 13).Value)) = "OUI")

    VerrouillerFormulaire False
End Sub

' --- Nouveau collaborateur ---
Private Sub cmdNouveau_Click()
    m_modeAjout = True
    m_ligneSelectionnee = 0
    ViderFormulaire
    VerrouillerFormulaire False
    txtNom.SetFocus
End Sub

' --- Supprimer ---
Private Sub cmdSupprimer_Click()
    If lstCollabs.ListIndex < 0 Then
        MsgBox "Sélectionnez d'abord un collaborateur.", vbExclamation: Exit Sub
    End If
    Dim nomSel As String: nomSel = txtNom.Text
    Dim rep As Integer
    rep = MsgBox("Supprimer " & nomSel & " ?", vbYesNo + vbWarning, "Confirmation")
    If rep = vbNo Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")
    ws.Rows(m_ligneSelectionnee).Delete

    ViderFormulaire
    VerrouillerFormulaire True
    m_ligneSelectionnee = 0
    ChargerListe
    MsgBox nomSel & " supprimé.", vbInformation
End Sub

' --- Sauvegarder ---
Private Sub cmdSauver_Click()
    ' Validations
    If Trim(txtNom.Text) = "" Then
        MsgBox "Le nom est obligatoire.", vbExclamation: txtNom.SetFocus: Exit Sub
    End If
    If Trim(cboProjet.Text) = "" Then
        MsgBox "Le projet est obligatoire.", vbExclamation: cboProjet.SetFocus: Exit Sub
    End If
    If chkConge.Value Then
        If Not IsDate(txtCongeD.Text) Or Not IsDate(txtCongeF.Text) Then
            MsgBox "Dates de congé invalides (format jj/mm/aaaa).", vbExclamation: Exit Sub
        End If
    End If
    If chkTT.Value Then
        If Not IsDate(txtTTD.Text) Or Not IsDate(txtTTF.Text) Then
            MsgBox "Dates TT invalides (format jj/mm/aaaa).", vbExclamation: Exit Sub
        End If
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Utilisateurs")
    Dim lr As Long

    If m_modeAjout Then
        ' Nouvelle ligne après la dernière
        lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    Else
        lr = m_ligneSelectionnee
    End If

    ws.Cells(lr, 1).Value = Trim(txtNom.Text)
    ws.Cells(lr, 2).Value = Trim(cboProjet.Text)
    ws.Cells(lr, 3).Value = Trim(txtVille.Text)
    ws.Cells(lr, 4).Value = Trim(txtZone.Text)

    ' Congé (col 5-7)
    ws.Cells(lr, 5).Value = IIf(chkConge.Value, "OUI", "NON")
    If chkConge.Value And IsDate(txtCongeD.Text) Then
        ws.Cells(lr, 6).Value = CDate(txtCongeD.Text)
        ws.Cells(lr, 6).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 6).Value = ""
    End If
    If chkConge.Value And IsDate(txtCongeF.Text) Then
        ws.Cells(lr, 7).Value = CDate(txtCongeF.Text)
        ws.Cells(lr, 7).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 7).Value = ""
    End If

    ' Transport (col 8)
    ws.Cells(lr, 8).Value = cboTransport.Text

    ' TT (col 9-11)
    ws.Cells(lr, 9).Value = IIf(chkTT.Value, "OUI", "NON")
    If chkTT.Value And IsDate(txtTTD.Text) Then
        ws.Cells(lr, 10).Value = CDate(txtTTD.Text)
        ws.Cells(lr, 10).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 10).Value = ""
    End If
    If chkTT.Value And IsDate(txtTTF.Text) Then
        ws.Cells(lr, 11).Value = CDate(txtTTF.Text)
        ws.Cells(lr, 11).NumberFormat = "dd/mm/yyyy"
    Else
        ws.Cells(lr, 11).Value = ""
    End If

    ' Renforts (col 12-13)
    ws.Cells(lr, 12).Value = IIf(chkRenfortP.Value, "OUI", "NON")
    ws.Cells(lr, 13).Value = IIf(chkRenfortI.Value, "OUI", "NON")

    ChargerListe
    VerrouillerFormulaire True
    m_modeAjout = False

    MsgBox "Collaborateur " & IIf(m_ligneSelectionnee = 0, "ajouté", "modifié") & " avec succès !", vbInformation
End Sub

Private Sub cmdAnnuler_Click()
    If m_ligneSelectionnee > 0 Then
        lstCollabs_Click
    Else
        ViderFormulaire
    End If
    m_modeAjout = False
    VerrouillerFormulaire True
End Sub

Private Sub cmdFermer_Click()
    Unload Me
End Sub

' --- Helpers ---
Private Sub ViderFormulaire()
    txtNom.Text = "": cboProjet.Text = "": txtVille.Text = "": txtZone.Text = ""
    chkConge.Value = False: txtCongeD.Text = "": txtCongeF.Text = ""
    cboTransport.Text = "NON"
    chkTT.Value = False: txtTTD.Text = "": txtTTF.Text = ""
    chkRenfortP.Value = False: chkRenfortI.Value = False
End Sub

Private Sub VerrouillerFormulaire(bVerrouille As Boolean)
    txtNom.Enabled = Not bVerrouille
    cboProjet.Enabled = Not bVerrouille
    txtVille.Enabled = Not bVerrouille
    txtZone.Enabled = Not bVerrouille
    chkConge.Enabled = Not bVerrouille
    txtCongeD.Enabled = Not bVerrouille
    txtCongeF.Enabled = Not bVerrouille
    cboTransport.Enabled = Not bVerrouille
    chkTT.Enabled = Not bVerrouille
    txtTTD.Enabled = Not bVerrouille
    txtTTF.Enabled = Not bVerrouille
    chkRenfortP.Enabled = Not bVerrouille
    chkRenfortI.Enabled = Not bVerrouille
    cmdSauver.Enabled = Not bVerrouille
    cmdAnnuler.Enabled = Not bVerrouille
End Sub

' Active/désactive les champs de date selon les checkboxes
Private Sub chkCONGE_Click()
    txtCongeD.Enabled = chkConge.Value
    txtCongeF.Enabled = chkConge.Value
End Sub

Private Sub chkTT_Click()
    txtTTD.Enabled = chkTT.Value
    txtTTF.Enabled = chkTT.Value
End Sub

' Ouverture rapide depuis le ruban / macro
Public Sub Ouvrir()
    UFUtilisateurs.Show
End Sub


