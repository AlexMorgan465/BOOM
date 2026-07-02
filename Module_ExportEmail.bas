Attribute VB_Name = "ExportEmail"
Option Explicit

' ============================================================
' MODULE : EXPORT + GENERATION EMAIL DE RENFORT
' A copier dans un NOUVEAU module standard du projet VBA
' (clic droit sur le projet > Insertion > Module, renommer
'  "ExportEmail", puis coller tout ce fichier).
' ============================================================


' ============================================================
' BOUTON 1 : EXPORTER UN CLASSEUR PAR ACTIVITE, DANS UN DOSSIER
'            NOMME COMME L'ACTIVITE
' ============================================================
' Résultat :
'   <dossier choisi>\AFEDIM\AFEDIM_S27_2026-07-02.xlsx
'   <dossier choisi>\EBRA\EBRA_S27_2026-07-02.xlsx
'   <dossier choisi>\GOOGLE LEADS\GOOGLE LEADS_S27_2026-07-02.xlsx
'   ... etc, un classeur par feuille d'activité (planning de semaine).
'
' Les feuilles "techniques" (Utilisateurs, BESOINS, ROTATION,
' CONSOLIDATION, PLANNING global) ne sont PAS exportées : seules les
' feuilles de planning par activité le sont. Si une nouvelle activité
' est ajoutée plus tard (nouvelle feuille), elle sera exportée
' automatiquement sans modifier ce code.
Sub ExporterPlanningsParActivite()
    Dim wsSource As Worksheet
    Dim wbExport As Workbook
    Dim dossierParent As String
    Dim dossierActivite As String
    Dim nomFichier As String
    Dim sem As Integer
    Dim feuillesTechniques As Variant
    Dim nomFeuille As String
    Dim compteur As Integer
    Dim fd As Object

    feuillesTechniques = Array("UTILISATEURS", "BESOINS", "ROTATION", "CONSOLIDATION", "PLANNING")

    ' Choix du dossier parent où créer les sous-dossiers par activité
    Set fd = Application.FileDialog(4) ' 4 = msoFileDialogFolderPicker
    fd.Title = "Choisir le dossier où exporter les plannings par activité"
    If fd.Show <> -1 Then Exit Sub ' annulé par l'utilisateur
    dossierParent = fd.SelectedItems(1)
    If Right(dossierParent, 1) <> "\" Then dossierParent = dossierParent & "\"

    On Error Resume Next
    sem = Application.WorksheetFunction.WeekNum(IIf(g_LundiCible = 0, LundiSemaineAuto(), g_LundiCible), 2)
    On Error GoTo 0

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    compteur = 0

    For Each wsSource In ThisWorkbook.Worksheets
        nomFeuille = wsSource.Name
        If Not EstFeuilleTechnique(nomFeuille, feuillesTechniques) Then

            ' Créer le sous-dossier au nom de l'activité s'il n'existe pas
            dossierActivite = dossierParent & NettoyerNomDossier(nomFeuille) & "\"
            If Dir(dossierActivite, vbDirectory) = "" Then MkDir dossierActivite

            ' Copier UNIQUEMENT cette feuille (planning de la semaine) dans un nouveau classeur
            wsSource.Copy
            Set wbExport = ActiveWorkbook

            ' Fige les formules en valeurs pour que l'export soit autonome
            With wbExport.Worksheets(1).UsedRange
                .Value = .Value
            End With

            nomFichier = dossierActivite & NettoyerNomDossier(nomFeuille) & "_S" & sem & "_" & Format(Now, "yyyy-mm-dd") & ".xlsx"
            wbExport.SaveAs Filename:=nomFichier, FileFormat:=xlOpenXMLWorkbook
            wbExport.Close SaveChanges:=False
            compteur = compteur + 1
        End If
    Next wsSource

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox compteur & " classeur(s) d'activité exporté(s) avec succès dans :" & Chr(10) & dossierParent, _
           vbInformation, "Export terminé"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Erreur lors de l'export : " & Err.Description, vbCritical, "Erreur"
End Sub

' Indique si une feuille est une feuille "technique" (à ne jamais exporter
' comme planning d'activité)
Private Function EstFeuilleTechnique(nom As String, liste As Variant) As Boolean
    Dim v As Variant
    For Each v In liste
        If UCase(nom) = v Then
            EstFeuilleTechnique = True
            Exit Function
        End If
    Next v
    EstFeuilleTechnique = False
End Function

' Nettoie un nom de feuille pour en faire un nom de dossier/fichier valide
' sous Windows (remplace les caractères interdits)
Private Function NettoyerNomDossier(nom As String) As String
    Dim car As Variant
    Dim resultat As String
    resultat = nom
    For Each car In Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        resultat = Replace(resultat, car, "-")
    Next car
    NettoyerNomDossier = Trim(resultat)
End Function


' ============================================================
' BOUTON 2 : GENERER L'EMAIL RECAP (sur le modele de l'exemple fourni)
' ============================================================
Sub GenererEmailRenfort()
    Dim olApp As Object
    Dim olMail As Object
    Dim corps As String
    Dim sem As Integer
    Dim lundi As Date, dimanche As Date

    lundi = IIf(g_LundiCible = 0, LundiSemaineAuto(), g_LundiCible)
    dimanche = lundi + 6
    sem = Application.WorksheetFunction.WeekNum(lundi, 2)

    corps = "Bonjour," & vbCrLf & vbCrLf
    corps = corps & "Vous trouverez ci-dessous la liste des collaborateurs qui seront en renfort durant la S" & sem & _
            " (" & Format(lundi, "dd/mm/yyyy") & " au " & Format(dimanche, "dd/mm/yyyy") & ") :" & vbCrLf & vbCrLf

    corps = corps & ConstruireBlocRenforts()

    corps = corps & vbCrLf & "Pour les congés payés :" & vbCrLf
    corps = corps & ConstruireBlocConges(lundi, dimanche)

    corps = corps & vbCrLf & "Congé maladie longue durée :" & vbCrLf
    corps = corps & ConstruireBlocMaladie()

    corps = corps & vbCrLf & "Fins de contrat :" & vbCrLf
    corps = corps & ConstruireBlocFinsContrat(lundi, dimanche)

    corps = corps & vbCrLf & "Départs :" & vbCrLf
    corps = corps & "[ A compléter manuellement - aucune colonne dédiée dans Utilisateurs ]" & vbCrLf

    corps = corps & vbCrLf & "Changement d'activité :" & vbCrLf
    corps = corps & "[ A compléter manuellement - aucune colonne dédiée dans Utilisateurs ]" & vbCrLf

    corps = corps & vbCrLf & "Cordialement," & vbCrLf

    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0

    If olApp Is Nothing Then
        MsgBox "Impossible de démarrer Outlook sur ce poste.", vbCritical
        Exit Sub
    End If

    Set olMail = olApp.CreateItem(0) ' 0 = olMailItem
    With olMail
        .Subject = "Renfort S" & sem & " : congés, congé maladie longue durée, départs, fins de contrat, changement d'activité"
        .Body = corps
        .Display   ' ouvre le brouillon pour relecture -> ne PAS remplacer par .Send
    End With

    MsgBox "Le brouillon d'email a été généré dans Outlook." & Chr(10) & _
           "Merci de compléter les sections manuelles (Départs / Changement d'activité) et de relire avant envoi.", _
           vbInformation, "Email généré"
End Sub


' ============================================================
' HELPERS - construction des blocs de texte de l'email
' ============================================================

' Bloc "renforts" jour par jour, à partir de BESOINS (colonnes 1=Projet,
' 3=Jour, 4=HeureDebut, 5=HeureFin, 7=Agents proposés, 9=Statut)
Private Function ConstruireBlocRenforts() As String
    Dim txt As String
    Dim wsB As Worksheet
    Dim collabs() As Collaborateur
    Dim nbCollab As Integer
    Dim r As Long, lastRow As Long
    Dim joursOrdre(1 To 7) As String
    Dim j As Integer

    joursOrdre(1) = "Lundi": joursOrdre(2) = "Mardi": joursOrdre(3) = "Mercredi"
    joursOrdre(4) = "Jeudi": joursOrdre(5) = "Vendredi": joursOrdre(6) = "Samedi": joursOrdre(7) = "Dimanche"

    If Not FeuilleExiste("BESOINS") Then
        ConstruireBlocRenforts = "(Feuille BESOINS introuvable)" & vbCrLf
        Exit Function
    End If

    nbCollab = LireCollaborateurs(collabs)

    Set wsB = ThisWorkbook.Sheets("BESOINS")
    lastRow = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row

    For j = 1 To 7
        Dim ligneJour As String: ligneJour = ""

        For r = 2 To lastRow
            Dim statut As String: statut = CStr(wsB.Cells(r, 9).Value)
            If Left(statut, 2) <> "OK" And Left(statut, 7) <> "PARTIEL" Then GoTo NextR
            If Trim(wsB.Cells(r, 3).Value) <> joursOrdre(j) Then GoTo NextR

            Dim projetBesoin As String: projetBesoin = Trim(wsB.Cells(r, 1).Value)
            Dim hdebut As String: hdebut = Trim(wsB.Cells(r, 4).Value)
            Dim hfin As String: hfin = Trim(wsB.Cells(r, 5).Value)
            Dim proposes As String: proposes = CStr(wsB.Cells(r, 7).Value)
            If proposes = "" Or proposes = "Aucun candidat disponible" Then GoTo NextR

            Dim agents() As String: agents = Split(proposes, " | ")
            Dim a As Integer
            For a = 0 To UBound(agents)
                Dim nomAgent As String: nomAgent = Trim(agents(a))
                If nomAgent = "" Then GoTo NextAgent

                Dim projetOrigine As String: projetOrigine = ""
                Dim ci As Integer
                For ci = 1 To nbCollab
                    If collabs(ci).nomComplet = nomAgent Then
                        projetOrigine = collabs(ci).projet
                        Exit For
                    End If
                Next ci

                ligneJour = ligneJour & "*  " & nomAgent & " (de " & hdebut & " à " & hfin & _
                            " sur " & projetBesoin & _
                            IIf(projetOrigine <> "", " / activité d'origine " & projetOrigine, "") & ")" & vbCrLf
NextAgent:
            Next a
NextR:
        Next r

        If ligneJour <> "" Then
            txt = txt & joursOrdre(j) & " :" & vbCrLf & ligneJour & vbCrLf
        End If
    Next j

    If txt = "" Then txt = "(Aucun renfort programmé cette semaine)" & vbCrLf
    ConstruireBlocRenforts = txt
End Function

' Bloc "congés payés" groupé par projet, sur la semaine [lundi ; dimanche]
Private Function ConstruireBlocConges(lundi As Date, dimanche As Date) As String
    Dim txt As String
    Dim collabs() As Collaborateur
    Dim nbCollab As Integer, i As Integer
    Dim projets As Object
    Set projets = CreateObject("Scripting.Dictionary")

    nbCollab = LireCollaborateurs(collabs)

    For i = 1 To nbCollab
        If collabs(i).EnConge Then
            If collabs(i).CongeFin >= lundi And collabs(i).CongeDebut <= dimanche Then
                Dim ligne As String
                ligne = "- " & collabs(i).nomComplet & " : du " & Format(collabs(i).CongeDebut, "dd/mm") & _
                        " au " & Format(collabs(i).CongeFin, "dd/mm") & " inclus"
                If collabs(i).TypeConge <> "" Then ligne = ligne & " (" & collabs(i).TypeConge & ")"

                If projets.Exists(collabs(i).projet) Then
                    projets(collabs(i).projet) = projets(collabs(i).projet) & ligne & vbCrLf
                Else
                    projets.Add collabs(i).projet, ligne & vbCrLf
                End If
            End If
        End If
    Next i

    Dim cle As Variant
    For Each cle In projets.Keys
        txt = txt & cle & " :" & vbCrLf & projets(cle) & vbCrLf
    Next cle

    If txt = "" Then txt = "(Aucun congé payé sur la période)" & vbCrLf
    ConstruireBlocConges = txt
End Function

' Bloc "congé maladie longue durée"
Private Function ConstruireBlocMaladie() As String
    Dim txt As String
    Dim collabs() As Collaborateur
    Dim nbCollab As Integer, i As Integer

    nbCollab = LireCollaborateurs(collabs)

    For i = 1 To nbCollab
        If collabs(i).EnMaladie Then
            txt = txt & "- " & collabs(i).nomComplet & " (" & collabs(i).projet & ") : congé maladie depuis le " & _
                  Format(collabs(i).DateArret, "dd/mm/yyyy")
            If collabs(i).DateReprise > collabs(i).DateArret Then
                txt = txt & " - reprise prévue le " & Format(collabs(i).DateReprise, "dd/mm/yyyy")
            End If
            txt = txt & vbCrLf
        End If
    Next i

    If txt = "" Then txt = "(Aucun cas en cours)" & vbCrLf
    ConstruireBlocMaladie = txt
End Function

' Bloc "fins de contrat" (CDD dont la date de sortie tombe dans les 14
' jours suivant le début de la semaine ciblée)
Private Function ConstruireBlocFinsContrat(lundi As Date, dimanche As Date) As String
    Dim txt As String
    Dim collabs() As Collaborateur
    Dim nbCollab As Integer, i As Integer
    Dim dSortie As Date

    nbCollab = LireCollaborateurs(collabs)

    For i = 1 To nbCollab
        If UCase(collabs(i).TypeContrat) = "CDD" And IsDate(collabs(i).DateSortie) Then
            dSortie = CDate(collabs(i).DateSortie)
            If dSortie >= lundi And dSortie <= dimanche + 14 Then
                txt = txt & "- " & collabs(i).nomComplet & " (" & collabs(i).projet & _
                      ") : fin de CDD le " & Format(dSortie, "dd/mm/yyyy") & vbCrLf
            End If
        End If
    Next i

    If txt = "" Then txt = "(Aucune fin de contrat prévue)" & vbCrLf
    ConstruireBlocFinsContrat = txt
End Function
