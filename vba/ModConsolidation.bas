Attribute VB_Name = "ModConsolidation"
Option Explicit

' Collecte des formulaires FANAF, vie et non-vie (montants en milliers de F CFA).
'
' Utilisation :
'   1. Lancer une fois InitialiserClasseur : cree la feuille Collecte, son bouton
'      Enregistrer (H2) et la cellule Code société (H5).
'   2. Copier le formulaire (colonnes A a F) et le coller en A1 de la feuille Collecte,
'      puis saisir le code societe en H5 (le meme code chaque annee pour une societe).
'   3. Cliquer sur Enregistrer : la saisie est controlee puis ajoutee a la base.
'
' Feuilles brutes (masquées), données telles que saisies :
'   Identification (brut), Emission&Prestations (brut), Chiffres clés (brut)
' Copies visibles, sans N° saisie, Date d'import ni Cellule source :
'   Identification, Emission&Prestations, Chiffres clés
'   Valeur (F CFA) = valeur brute (milliers F CFA) * 1000, par formule
'
' Chaque table commence par une colonne Clé, unique par ligne, pour RECHERCHEV :
'   Identification       : Code|Annee
'   Emission&Prestations : Code|Bloc|Categorie|Rubrique|Mesure|Annee
'   Chiffres clés        : Code|Rubrique|Annee
' Le code societe (ex. BJ-AFGV) reste stable quand la raison sociale change.
' Les ratios et agregats ne sont pas stockes : ils se calculent dans les etudes.
'
' Lecture du formulaire, sans adresses fixes :
'   - les blocs sont reperes par leurs titres en colonne B :
'       A - EMISSIONS, B - PRESTATIONS (vie) ou B - SINISTRALITE (non-vie), C - AUTRES CHIFFRES
'   - la ligne "Branches" porte les annees : N-1 en colonne B, N en colonne D
'   - blocs A et B :
'       vie     : un montant par annee (N-1 en B, N en D)
'       non-vie : une ligne de sous-en-tetes donne deux mesures par annee
'                 (N-1 en B et C, N en D et E) ; les ratios ("CS / PA") sont ignores
'       les lignes "Assurances ..." donnent la categorie ; Total et Ensemble sont ignores
'   - bloc C : lu ligne a ligne jusqu'a "Taux de couverture" ;
'       Autres actifs et Taux de couverture (formules) sont ignores
'
' Fichier encodé en Windows-1252 (ANSI), l'encodage attendu par l'éditeur VBA.

Private Const F_COLLECTE As String = "Collecte"
Private Const F_ID As String = "Identification"
Private Const F_EP As String = "Emission&Prestations"
Private Const SUFFIXE_BRUT As String = " (brut)"
Private Const NOM_BOUTON As String = "btnEnregistrer"
Private Const CELLULE_CODE As String = "H5"   ' code societe, hors du bloc colle (A:F)

' Derniere ligne examinee sur la feuille Collecte.
Private Const LIGNE_MAX As Long = 150

' Tables de destination d'une entree lue
Private Const T_EP As Long = 2
Private Const T_CC As Long = 3

' Position des blocs dans la feuille Collecte, renseignee par Reperer.
Private mA As Long      ' ligne du titre "A - EMISSIONS NETTES"
Private mB As Long      ' ligne du titre "B - PRESTATIONS VERSEES" ou "B - SINISTRALITE"
Private mC As Long      ' ligne du titre "C - AUTRES CHIFFRES CLES"
Private mFinC As Long   ' derniere ligne du bloc C ("Taux de couverture")

Private Function NomChiffresCles() As String
    NomChiffresCles = "Chiffres clés"
End Function

' ---------------------------------------------------------------------------------
' Mise en place : feuille Collecte + bouton Enregistrer
' ---------------------------------------------------------------------------------
Public Sub InitialiserClasseur()
    Dim ws As Worksheet, cible As Range, btn As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(F_COLLECTE)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = F_COLLECTE
    End If

    On Error Resume Next
    ws.Buttons(NOM_BOUTON).Delete
    On Error GoTo 0

    Set cible = ws.Range("H2")
    Set btn = ws.Buttons.Add(cible.Left, cible.Top, 120, 28)
    btn.Name = NOM_BOUTON
    btn.Caption = "Enregistrer"
    btn.OnAction = "EnregistrerCollecte"

    ws.Range(CELLULE_CODE).Offset(-1, 0).Value = "Code société"
    ws.Range(CELLULE_CODE).Offset(-1, 0).Font.Bold = True
    ws.Range(CELLULE_CODE).NumberFormat = "@"
    ws.Range(CELLULE_CODE).BorderAround xlContinuous, xlThin

    FeuilleBrute F_ID & SUFFIXE_BRUT, EntetesIdentification()
    FeuilleBrute F_EP & SUFFIXE_BRUT, EntetesEmissionsPrestations()
    FeuilleBrute NomChiffresCles() & SUFFIXE_BRUT, EntetesChiffresCles()
    RafraichirCopies
    MasquerBrutes

    ws.Activate
    MsgBox "Feuille Collecte prête : collez le formulaire en A1, saisissez le code société en " & _
        CELLULE_CODE & " puis cliquez sur Enregistrer.", vbInformation, "Collecte"
End Sub

' ---------------------------------------------------------------------------------
' Bouton Enregistrer
' ---------------------------------------------------------------------------------
Public Sub EnregistrerCollecte()
    Dim ws As Worksheet, wsIdB As Worksheet, wsEpB As Worksheet, wsCcB As Worksheet
    Dim erreurs As String, societe As String, code As String, idSaisie As Long, anN As Long
    Dim anciens As Collection, entrees As Collection
    Dim numErr As Long, descErr As String

    Set ws = ThisWorkbook.Worksheets(F_COLLECTE)

    erreurs = ControlerCollecte(ws)
    If Len(erreurs) > 0 Then
        MsgBox "Enregistrement refusé :" & vbLf & erreurs, vbExclamation, "Collecte"
        Exit Sub
    End If

    societe = Texte(ws.Range("B1").Value)
    code = CodeSociete(ws)
    anN = AnneeN1(ws) + 1

    Set wsIdB = FeuilleBrute(F_ID & SUFFIXE_BRUT, EntetesIdentification())
    Set wsEpB = FeuilleBrute(F_EP & SUFFIXE_BRUT, EntetesEmissionsPrestations())
    Set wsCcB = FeuilleBrute(NomChiffresCles() & SUFFIXE_BRUT, EntetesChiffresCles())

    ' Meme code societe et meme annee deja enregistres : remplacement sur confirmation
    Set anciens = SaisiesExistantes(wsIdB, code, anN)
    If anciens.Count > 0 Then
        If MsgBox(code & " (" & anN & ") est déjà enregistré." & vbLf & _
                  "Remplacer les données existantes ?", vbYesNo + vbQuestion, "Collecte") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Erreur

    If anciens.Count > 0 Then
        SupprimerSaisies wsIdB, anciens
        SupprimerSaisies wsEpB, anciens
        SupprimerSaisies wsCcB, anciens
    End If

    idSaisie = Application.WorksheetFunction.Max(wsIdB.Columns(1)) + 1
    Set entrees = LireEntrees(ws)
    EcrireIdentification ws, idSaisie, wsIdB
    EcrireEntrees ws, idSaisie, entrees, wsEpB, wsCcB

    RafraichirCopies
    MasquerBrutes
    ws.Activate
    Application.ScreenUpdating = True

    If MsgBox(societe & " (" & code & ", " & anN & ") enregistrée." & vbLf & vbLf & _
              "Vider la feuille Collecte pour la saisie suivante ?", vbYesNo + vbInformation, "Collecte") = vbYes Then
        ws.Range("A1:F" & LIGNE_MAX).ClearContents
        ws.Range(CELLULE_CODE).ClearContents
    End If
    Exit Sub

Erreur:
    numErr = Err.Number
    descErr = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Erreur " & numErr & " :" & vbLf & descErr, vbCritical, "Collecte"
End Sub

' ---------------------------------------------------------------------------------
' Reperage des blocs
' ---------------------------------------------------------------------------------

' Repere les titres des blocs et la fin du bloc C. Renvoie False si la mise en page
' ne correspond pas a un formulaire FANAF.
Private Function Reperer(ByVal ws As Worksheet) As Boolean
    Dim r As Long, t As String

    mA = 0: mB = 0: mC = 0: mFinC = 0
    For r = 1 To LIGNE_MAX
        t = Texte(ws.Cells(r, 2).Value)
        If mA = 0 And InStr(1, t, "EMISSIONS", vbTextCompare) > 0 Then mA = r
        If mB = 0 And (InStr(1, t, "PRESTATIONS", vbTextCompare) > 0 _
                    Or InStr(1, t, "SINISTRALITE", vbTextCompare) > 0) Then mB = r
        If mC = 0 And InStr(1, t, "AUTRES CHIFFRES", vbTextCompare) > 0 Then mC = r
    Next r
    If mA = 0 Or mB = 0 Or mC = 0 Then Exit Function

    ' Bloc C : de la ligne apres "Rubriques" jusqu'a "Taux de couverture" (ou 1re ligne vide)
    For r = mC + 2 To LIGNE_MAX
        t = Texte(ws.Cells(r, 1).Value)
        If Len(t) = 0 Then Exit For
        mFinC = r
        If InStr(1, t, "Taux de couverture", vbTextCompare) > 0 Then Exit For
    Next r

    Reperer = UCase$(Texte(ws.Range("A1").Value)) = "NOM DE LA SOCIETE" _
        And mA < mB And mB < mC And mFinC > mC + 1 _
        And UCase$(Texte(ws.Cells(mA + 1, 1).Value)) = "BRANCHES" _
        And UCase$(Texte(ws.Cells(mB + 1, 1).Value)) = "BRANCHES" _
        And UCase$(Texte(ws.Cells(mC + 1, 1).Value)) = "RUBRIQUES"
End Function

' "Vie" ou "Non-vie" selon le titre du bloc B.
Private Function TypeSociete(ByVal ws As Worksheet) As String
    If InStr(1, Texte(ws.Cells(mB, 2).Value), "SINISTRALITE", vbTextCompare) > 0 Then
        TypeSociete = "Non-vie"
    Else
        TypeSociete = "Vie"
    End If
End Function

Private Function NomBloc(ByVal titre As String) As String
    If InStr(1, titre, "EMISSIONS", vbTextCompare) > 0 Then
        NomBloc = "Émissions"
    ElseIf InStr(1, titre, "PRESTATIONS", vbTextCompare) > 0 Then
        NomBloc = "Prestations"
    ElseIf InStr(1, titre, "SINISTRALITE", vbTextCompare) > 0 Then
        NomBloc = "Sinistralité"
    Else
        NomBloc = titre
    End If
End Function

' Mesure unique d'un bloc vie, d'apres son titre.
Private Function MesureParDefaut(ByVal titre As String) As String
    If InStr(1, titre, "EMISSIONS", vbTextCompare) > 0 Then
        MesureParDefaut = "Émissions nettes"
    ElseIf InStr(1, titre, "PRESTATIONS", vbTextCompare) > 0 Then
        MesureParDefaut = "Prestations versées"
    Else
        MesureParDefaut = "Montant"
    End If
End Function

' Libelle de mesure lu dans la ligne de sous-en-tetes d'un bloc non-vie.
Private Function NomMesure(ByVal v As Variant) As String
    Dim s As String
    s = Texte(v)
    If StrComp(s, "(CS)", vbTextCompare) = 0 Then s = "Charges de sinistres (CS)"
    NomMesure = s
End Function

' Code societe saisi en H5, en majuscules et sans espaces superflus.
Private Function CodeSociete(ByVal ws As Worksheet) As String
    CodeSociete = UCase$(Texte(ws.Range(CELLULE_CODE).Value))
End Function

Private Function AnneeN1(ByVal ws As Worksheet) As Long
    AnneeN1 = CLng(Nombre(ws.Cells(mA + 1, 2).Value))
End Function

' Rubriques du bloc C calculees par formule dans le formulaire : non reprises.
Private Function EstDerivee(ByVal rubrique As String) As Boolean
    EstDerivee = StrComp(rubrique, "Autres actifs", vbTextCompare) = 0 _
        Or InStr(1, rubrique, "Taux de couverture", vbTextCompare) > 0
End Function

' ---------------------------------------------------------------------------------
' Entrees du formulaire : une par valeur saisie
'   Array(table, bloc, categorie, rubrique, mesure, annee, cellule)
' ---------------------------------------------------------------------------------
Private Function LireEntrees(ByVal ws As Worksheet) As Collection
    Dim res As New Collection, r As Long, anN1 As Long, rubrique As String

    anN1 = AnneeN1(ws)
    AjouterEntreesBloc ws, mA, mB - 1, anN1, res
    AjouterEntreesBloc ws, mB, mC - 1, anN1, res

    For r = mC + 2 To mFinC
        rubrique = Libelle(ws.Cells(r, 1).Value)
        If Not EstDerivee(rubrique) Then
            res.Add Array(T_CC, "", "", rubrique, "", anN1, ws.Cells(r, 3))
            res.Add Array(T_CC, "", "", rubrique, "", anN1 + 1, ws.Cells(r, 4))
        End If
    Next r
    Set LireEntrees = res
End Function

' Bloc A ou B, du titre (ligne t) jusqu'a la ligne fin.
Private Sub AjouterEntreesBloc(ByVal ws As Worksheet, ByVal t As Long, ByVal fin As Long, _
                               ByVal anN1 As Long, ByVal res As Collection)
    Dim bloc As String, categorie As String, rubrique As String, titre As String
    Dim mesures As New Collection, m As Variant
    Dim debut As Long, r As Long, c As Long, vide As Boolean

    titre = Texte(ws.Cells(t, 2).Value)
    bloc = NomBloc(titre)

    ' Sous-en-tetes (non-vie) : colonne A vide et libelle texte en colonne B
    If Len(Texte(ws.Cells(t + 2, 1).Value)) = 0 And Len(Texte(ws.Cells(t + 2, 2).Value)) > 0 _
       And Not IsNumeric(ws.Cells(t + 2, 2).Value) Then
        ' mesure 1 : N-1 en B, N en D ; mesure 2 : N-1 en C, N en E ; ratios ignores
        If InStr(ws.Cells(t + 2, 2).Value, "/") = 0 Then mesures.Add Array(NomMesure(ws.Cells(t + 2, 2).Value), 2, 4)
        If InStr(ws.Cells(t + 2, 3).Value, "/") = 0 Then mesures.Add Array(NomMesure(ws.Cells(t + 2, 3).Value), 3, 5)
        debut = t + 3
        categorie = "Affaires directes"
    Else
        mesures.Add Array(MesureParDefaut(titre), 2, 4)
        debut = t + 2
        categorie = ""
    End If

    For r = debut To fin
        rubrique = Libelle(ws.Cells(r, 1).Value)
        If Len(rubrique) = 0 Then GoTo Suivante

        vide = True
        For c = 2 To 6
            If Len(Texte(ws.Cells(r, c).Value)) > 0 Then vide = False
        Next c

        If vide And StrComp(Left$(rubrique, 11), "Assurances ", vbTextCompare) = 0 Then
            ' Ligne de categorie : "Assurances individuelles" -> "Individuelles"
            categorie = Mid$(rubrique, 12)
            If Len(categorie) > 0 Then categorie = UCase$(Left$(categorie, 1)) & Mid$(categorie, 2)
        ElseIf StrComp(Left$(rubrique, 5), "Total", vbTextCompare) = 0 _
            Or StrComp(rubrique, "Ensemble", vbTextCompare) = 0 Then
            ' Agregats : recalcules dans les etudes
        Else
            For Each m In mesures
                If StrComp(rubrique, "Acceptations", vbTextCompare) = 0 Then
                    res.Add Array(T_EP, bloc, "Acceptations", rubrique, m(0), anN1, ws.Cells(r, m(1)))
                    res.Add Array(T_EP, bloc, "Acceptations", rubrique, m(0), anN1 + 1, ws.Cells(r, m(2)))
                Else
                    res.Add Array(T_EP, bloc, categorie, rubrique, m(0), anN1, ws.Cells(r, m(1)))
                    res.Add Array(T_EP, bloc, categorie, rubrique, m(0), anN1 + 1, ws.Cells(r, m(2)))
                End If
            Next m
        End If
Suivante:
    Next r
End Sub

' ---------------------------------------------------------------------------------
' Controles de la feuille Collecte. Renvoie "" si tout est correct.
' ---------------------------------------------------------------------------------
Private Function ControlerCollecte(ByVal ws As Worksheet) As String
    Dim msg As String, invalides As String, v As Variant, e As Variant, c As Range
    Dim entrees As Collection

    If Not Reperer(ws) Then
        ControlerCollecte = "- Le bloc collé en A1 n'a pas la mise en page du formulaire" & vbLf & _
            "  (titres A - EMISSIONS, B - PRESTATIONS ou SINISTRALITE et C - AUTRES CHIFFRES" & vbLf & _
            "  introuvables, ou lignes Branches / Rubriques absentes)."
        Exit Function
    End If

    ' Champs obligatoires
    If Len(Texte(ws.Range("B1").Value)) = 0 Then msg = msg & vbLf & "- Nom de la société vide (B1)."
    If Len(Texte(ws.Range("B2").Value)) = 0 Then msg = msg & vbLf & "- Pays vide (B2)."
    If Len(CodeSociete(ws)) = 0 Then msg = msg & vbLf & "- Code société vide (" & CELLULE_CODE & ")."
    v = Nombre(ws.Cells(mA + 1, 2).Value)
    If Not IsNumeric(v) Or IsEmpty(v) Then
        msg = msg & vbLf & "- Année N-1 invalide (" & ws.Cells(mA + 1, 2).Address(False, False) & ")."
        ControlerCollecte = Mid$(msg, 2)
        Exit Function
    ElseIf v < 1990 Or v > 2100 Or v <> Int(v) Then
        msg = msg & vbLf & "- Année N-1 invalide (" & ws.Cells(mA + 1, 2).Address(False, False) & ")."
        ControlerCollecte = Mid$(msg, 2)
        Exit Function
    End If

    ' Date de creation : vide ou date
    v = ws.Range("B4").Value
    If IsError(v) Then
        msg = msg & vbLf & "- Date de création invalide (B4)."
    ElseIf Not IsEmpty(v) And VarType(v) <> vbDate And Not IsNumeric(v) And Not IsDate(v) Then
        msg = msg & vbLf & "- Date de création invalide (B4)."
    End If

    ' Effectifs : vides ou numeriques
    For Each c In ws.Range("D6:D8").Cells
        If Not EstNombreOuVide(c.Value) Then invalides = invalides & " " & c.Address(False, False)
    Next c

    ' Cellules chiffrees des blocs : vides ou numeriques (nombres colles en texte acceptes)
    Set entrees = LireEntrees(ws)
    If entrees.Count = 0 Then msg = msg & vbLf & "- Aucune donnée trouvée dans les blocs A, B et C."
    For Each e In entrees
        Set c = e(6)
        If Not EstNombreOuVide(c.Value) Then invalides = invalides & " " & c.Address(False, False)
    Next e
    If Len(invalides) > 0 Then msg = msg & vbLf & "- Valeurs non numériques :" & invalides

    If Len(msg) > 0 Then msg = Mid$(msg, 2)
    ControlerCollecte = msg
End Function

Private Function EstNombreOuVide(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    v = Nombre(v)
    EstNombreOuVide = IsEmpty(v) Or IsNumeric(v)
End Function

' ---------------------------------------------------------------------------------
' Ecriture dans les feuilles brutes
' ---------------------------------------------------------------------------------
Private Function EntetesIdentification() As Variant
    EntetesIdentification = Array("N° saisie", "Clé", "Code société", "Société", "Pays", "Type", _
        "Directeur général", "Date de création", "Capital social (F CFA)", "Cadres", "Maîtrise", _
        "Employés", "Année N", "Date d'import")
End Function

Private Function EntetesEmissionsPrestations() As Variant
    EntetesEmissionsPrestations = Array("N° saisie", "Clé", "Code société", "Société", "Pays", "Type", _
        "Bloc", "Catégorie", "Rubrique", "Mesure", "Année", "Valeur (milliers F CFA)", "Cellule source")
End Function

Private Function EntetesChiffresCles() As Variant
    EntetesChiffresCles = Array("N° saisie", "Clé", "Code société", "Société", "Pays", "Type", _
        "Rubrique", "Année", "Valeur (milliers F CFA)", "Cellule source")
End Function

' Une ligne par saisie.
Private Sub EcrireIdentification(ByVal ws As Worksheet, ByVal idSaisie As Long, ByVal dest As Worksheet)
    AjouterLigne dest, Array(idSaisie, _
        Cle(CodeSociete(ws), AnneeN1(ws) + 1), _
        CodeSociete(ws), _
        Texte(ws.Range("B1").Value), _
        Texte(ws.Range("B2").Value), _
        TypeSociete(ws), _
        Texte(ws.Range("B3").Value), _
        ValeurDate(ws.Range("B4").Value), _
        ValeurCapital(ws.Range("B5").Value), _
        Nombre(ws.Range("D6").Value), _
        Nombre(ws.Range("D7").Value), _
        Nombre(ws.Range("D8").Value), _
        AnneeN1(ws) + 1, _
        Now)
End Sub

' Une ligne par entree, dans Emission&Prestations ou Chiffres clés.
Private Sub EcrireEntrees(ByVal ws As Worksheet, ByVal idSaisie As Long, ByVal entrees As Collection, _
                          ByVal wsEpB As Worksheet, ByVal wsCcB As Worksheet)
    Dim e As Variant, c As Range
    Dim code As String, societe As String, pays As String, typ As String

    code = CodeSociete(ws)
    societe = Texte(ws.Range("B1").Value)
    pays = Texte(ws.Range("B2").Value)
    typ = TypeSociete(ws)

    For Each e In entrees
        Set c = e(6)
        If e(0) = T_EP Then
            AjouterLigne wsEpB, Array(idSaisie, Cle(code, e(1), e(2), e(3), e(4), e(5)), _
                code, societe, pays, typ, e(1), e(2), e(3), e(4), e(5), _
                Nombre(c.Value), c.Address(False, False))
        Else
            AjouterLigne wsCcB, Array(idSaisie, Cle(code, e(3), e(5)), _
                code, societe, pays, typ, e(3), e(5), _
                Nombre(c.Value), c.Address(False, False))
        End If
    Next e
End Sub

Private Sub AjouterLigne(ByVal dest As Worksheet, ByVal valeurs As Variant)
    Dim lig As Long
    lig = dest.Cells(dest.Rows.Count, 1).End(xlUp).Row + 1
    dest.Cells(lig, 1).Resize(1, UBound(valeurs) - LBound(valeurs) + 1).Value = valeurs
End Sub

' ---------------------------------------------------------------------------------
' Doublons : meme code societe et meme annee N
' ---------------------------------------------------------------------------------
Private Function SaisiesExistantes(ByVal wsIdB As Worksheet, ByVal code As String, ByVal anN As Long) As Collection
    Dim res As New Collection, r As Long, der As Long

    der = wsIdB.Cells(wsIdB.Rows.Count, 1).End(xlUp).Row
    For r = 2 To der
        If StrComp(Texte(wsIdB.Cells(r, 3).Value), code, vbTextCompare) = 0 _
           And Val(wsIdB.Cells(r, 13).Value) = anN Then
            res.Add wsIdB.Cells(r, 1).Value
        End If
    Next r
    Set SaisiesExistantes = res
End Function

Private Sub SupprimerSaisies(ByVal ws As Worksheet, ByVal ids As Collection)
    Dim r As Long, id As Variant

    For r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        For Each id In ids
            If ws.Cells(r, 1).Value = id Then
                ws.Rows(r).Delete
                Exit For
            End If
        Next id
    Next r
End Sub

' ---------------------------------------------------------------------------------
' Feuilles brutes et copies visibles
' ---------------------------------------------------------------------------------

' Renvoie la feuille brute (creee avec ses en-tetes si absente), sans la vider.
Private Function FeuilleBrute(ByVal nom As String, ByVal entetes As Variant) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
        With ws.Range("A1").Resize(1, UBound(entetes) - LBound(entetes) + 1)
            .Value = entetes
            .Font.Bold = True
        End With
    End If
    Set FeuilleBrute = ws
End Function

Private Sub MasquerBrutes()
    ThisWorkbook.Worksheets(F_ID & SUFFIXE_BRUT).Visible = xlSheetHidden
    ThisWorkbook.Worksheets(F_EP & SUFFIXE_BRUT).Visible = xlSheetHidden
    ThisWorkbook.Worksheets(NomChiffresCles() & SUFFIXE_BRUT).Visible = xlSheetHidden
End Sub

' Recree les trois copies visibles a partir des feuilles brutes.
Public Sub RafraichirCopies()
    Dim wsIdB As Worksheet, wsEpB As Worksheet, wsCcB As Worksheet

    Set wsIdB = FeuilleBrute(F_ID & SUFFIXE_BRUT, EntetesIdentification())
    Set wsEpB = FeuilleBrute(F_EP & SUFFIXE_BRUT, EntetesEmissionsPrestations())
    Set wsCcB = FeuilleBrute(NomChiffresCles() & SUFFIXE_BRUT, EntetesChiffresCles())

    wsIdB.Columns("H").NumberFormat = "dd/mm/yyyy"
    wsIdB.Columns("I").NumberFormat = "#,##0"
    wsIdB.Columns("N").NumberFormat = "dd/mm/yyyy hh:mm"
    wsEpB.Columns("L").NumberFormat = "#,##0.000"
    wsCcB.Columns("I").NumberFormat = "#,##0.000"

    ' On saute la colonne A (N° saisie) et on s'arrête avant Date d'import / Cellule source ;
    ' la valeur est reprise par formule * 1000.
    CreerCopie wsIdB, F_ID, Array("Clé", "Code société", "Société", "Pays", "Type", "Directeur général", _
        "Date de création", "Capital social (F CFA)", "Cadres", "Maîtrise", "Employés", "Année N"), 12, ""
    CreerCopie wsEpB, F_EP, Array("Clé", "Code société", "Société", "Pays", "Type", "Bloc", "Catégorie", _
        "Rubrique", "Mesure", "Année", "Valeur (F CFA)"), 10, "L"
    CreerCopie wsCcB, NomChiffresCles(), Array("Clé", "Code société", "Société", "Pays", "Type", _
        "Rubrique", "Année", "Valeur (F CFA)"), 7, "I"
End Sub

' Copie visible d'une feuille brute : colonnes B a (1 + nbCols) en valeurs, puis,
' si colValeur est renseignée, une colonne Valeur (F CFA) = brute!colValeur * 1000.
Private Sub CreerCopie(ByVal wsBrut As Worksheet, ByVal nom As String, ByVal entetes As Variant, _
                       ByVal nbCols As Long, ByVal colValeur As String)
    Dim ws As Worksheet, nbEntetes As Long, derLig As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
    End If

    ws.Cells.Clear
    nbEntetes = UBound(entetes) - LBound(entetes) + 1
    With ws.Range("A1").Resize(1, nbEntetes)
        .Value = entetes
        .Font.Bold = True
    End With

    derLig = wsBrut.Cells(wsBrut.Rows.Count, 1).End(xlUp).Row
    If derLig >= 2 Then
        ws.Range("A2").Resize(derLig - 1, nbCols).Value = wsBrut.Range("B2").Resize(derLig - 1, nbCols).Value
        If Len(colValeur) > 0 Then
            With ws.Cells(2, nbCols + 1).Resize(derLig - 1, 1)
                .Formula = "=IF('" & wsBrut.Name & "'!" & colValeur & "2="""",""""," & _
                    "'" & wsBrut.Name & "'!" & colValeur & "2*1000)"
                .NumberFormat = "#,##0"
            End With
        End If
    End If

    If nom = F_ID Then
        ws.Columns("G").NumberFormat = "dd/mm/yyyy"
        ws.Columns("H").NumberFormat = "#,##0"
    End If
    ws.Visible = xlSheetVisible
    ws.Range("A1").Resize(1, nbEntetes).EntireColumn.AutoFit
End Sub

' ---------------------------------------------------------------------------------
' Conversions
' ---------------------------------------------------------------------------------

' Cle de recherche : elements separes par "|", ex. "BJ-AFGV|Marge disponible|2024".
Private Function Cle(ParamArray parts() As Variant) As String
    Dim i As Long, s As String
    For i = LBound(parts) To UBound(parts)
        If i > LBound(parts) Then s = s & "|"
        s = s & CStr(parts(i))
    Next i
    Cle = s
End Function

Private Function Texte(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    Texte = Application.WorksheetFunction.Trim(CStr(v))
End Function

' Libelle de rubrique nettoye (espaces, fautes de frappe des formulaires, majuscule initiale).
Private Function Libelle(ByVal v As Variant) As String
    Dim s As String
    s = Texte(v)
    s = Replace(s, "autes actifs", "autres actifs", , , vbTextCompare)
    s = Replace(s, "propres er ", "propres et ", , , vbTextCompare)
    If Len(s) > 0 Then s = UCase$(Left$(s, 1)) & Mid$(s, 2)
    Libelle = s
End Function

' Cellule vide ou en erreur -> Empty (distinct d'un vrai 0).
' Un nombre colle en texte ("1 177 646", espaces insecables compris) est converti.
Private Function Nombre(ByVal v As Variant) As Variant
    Dim s As String

    If IsError(v) Then
        Nombre = Empty
    ElseIf IsEmpty(v) Then
        Nombre = Empty
    ElseIf VarType(v) = vbString Then
        s = Replace(Replace(Replace(v, " ", ""), ChrW(160), ""), ChrW(8239), "")
        If Len(s) = 0 Then
            Nombre = Empty
        ElseIf IsNumeric(s) Then
            Nombre = CDbl(s)
        Else
            Nombre = v
        End If
    ElseIf IsNumeric(v) Then
        Nombre = CDbl(v)
    Else
        Nombre = v
    End If
End Function

Private Function ValeurDate(ByVal v As Variant) As Variant
    If IsError(v) Then
        ValeurDate = Empty
    ElseIf IsEmpty(v) Then
        ValeurDate = Empty
    ElseIf VarType(v) = vbDate Then
        ValeurDate = v
    ElseIf IsNumeric(v) Then
        ValeurDate = CDate(CDbl(v))
    ElseIf IsDate(v) Then
        ValeurDate = CDate(v)
    Else
        ValeurDate = Texte(v)
    End If
End Function

' "3 299 350 000 F CFA" -> 3299350000 (le capital est deja en F CFA, pas en milliers)
Private Function ValeurCapital(ByVal v As Variant) As Variant
    Dim s As String, c As String, chiffres As String, i As Long

    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        ValeurCapital = CDbl(v)
        Exit Function
    End If

    s = CStr(v)
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If c Like "#" Then chiffres = chiffres & c
    Next i
    If Len(chiffres) > 0 Then ValeurCapital = CDbl(chiffres)
End Function
