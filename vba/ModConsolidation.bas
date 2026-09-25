Attribute VB_Name = "ModConsolidation"
Option Explicit

' Collecte des formulaires au format "Modele.xlsx" (montants en milliers de F CFA).
'
' Utilisation :
'   1. Lancer une fois InitialiserClasseur : cree la feuille Collecte et son bouton Enregistrer.
'   2. Copier le bloc A1:F66 d'un formulaire et le coller en A1 de la feuille Collecte.
'   3. Cliquer sur Enregistrer : la saisie est controlee puis ajoutee a la base.
'
' Feuilles brutes (masquees), donnees telles que saisies :
'   Identification_brut, Emission&Prestations_brut, Chiffres cles_brut
' Copies visibles, sans Id_saisie, Date_import ni Cellule_source :
'   Identification, Emission&Prestations, Chiffres cles
'   Valeur_FCFA = valeur brute (milliers) * 1000, par formule
'
' Chaque table commence par une colonne Cle, unique par ligne, pour RECHERCHEV :
'   Identification       : Pays|Societe|Annee
'   Emission&Prestations : Pays|Societe|Bloc|Categorie|Rubrique|Annee
'   Chiffres cles        : Pays|Societe|Rubrique|Annee
' Les ratios et agregats ne sont pas stockes : ils se calculent dans les etudes.
'
' Le code source est sans accents pour s'importer sans souci dans l'editeur VBA ;
' le "e accent aigu" du nom de feuille est produit par ChrW(233).

Private Const F_COLLECTE As String = "Collecte"
Private Const F_ID As String = "Identification"
Private Const F_EP As String = "Emission&Prestations"
Private Const SUFFIXE_BRUT As String = "_brut"
Private Const NOM_BOUTON As String = "btnEnregistrer"

' Cellules de saisie numerique du formulaire (les autres cellules chiffrees sont des formules).
Private Const CELLULES_NUMERIQUES As String = _
    "D6:D8,B13:B18,D13:D18,B20:B25,D20:D25,B27,D27," & _
    "B33:B38,D33:D38,B40:B45,D40:D45,B47,D47,C51:D64"

Private Function NomChiffresCles() As String
    NomChiffresCles = "Chiffres cl" & ChrW(233) & "s"
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

    FeuilleBrute F_ID & SUFFIXE_BRUT, EntetesIdentification()
    FeuilleBrute F_EP & SUFFIXE_BRUT, EntetesEmissionsPrestations()
    FeuilleBrute NomChiffresCles() & SUFFIXE_BRUT, EntetesChiffresCles()
    RafraichirCopies
    MasquerBrutes

    ws.Activate
    MsgBox "Feuille Collecte prete : collez le formulaire en A1 puis cliquez sur Enregistrer.", _
        vbInformation, "Collecte"
End Sub

' ---------------------------------------------------------------------------------
' Bouton Enregistrer
' ---------------------------------------------------------------------------------
Public Sub EnregistrerCollecte()
    Dim ws As Worksheet, wsIdB As Worksheet, wsEpB As Worksheet, wsCcB As Worksheet
    Dim erreurs As String, societe As String, idSaisie As Long, anN As Long
    Dim anciens As Collection
    Dim numErr As Long, descErr As String

    Set ws = ThisWorkbook.Worksheets(F_COLLECTE)

    erreurs = ControlerCollecte(ws)
    If Len(erreurs) > 0 Then
        MsgBox "Enregistrement refuse :" & vbLf & erreurs, vbExclamation, "Collecte"
        Exit Sub
    End If

    societe = Texte(ws.Range("B1").Value)
    anN = CLng(ws.Range("B10").Value) + 1

    Set wsIdB = FeuilleBrute(F_ID & SUFFIXE_BRUT, EntetesIdentification())
    Set wsEpB = FeuilleBrute(F_EP & SUFFIXE_BRUT, EntetesEmissionsPrestations())
    Set wsCcB = FeuilleBrute(NomChiffresCles() & SUFFIXE_BRUT, EntetesChiffresCles())

    ' Meme societe et meme annee deja enregistrees : remplacement sur confirmation
    Set anciens = SaisiesExistantes(wsIdB, societe, anN)
    If anciens.Count > 0 Then
        If MsgBox(societe & " (" & anN & ") est deja enregistree." & vbLf & _
                  "Remplacer les donnees existantes ?", vbYesNo + vbQuestion, "Collecte") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo Erreur

    If anciens.Count > 0 Then
        SupprimerSaisies wsIdB, anciens
        SupprimerSaisies wsEpB, anciens
        SupprimerSaisies wsCcB, anciens
    End If

    idSaisie = Application.WorksheetFunction.Max(wsIdB.Columns(1)) + 1
    LireIdentification ws, idSaisie, wsIdB
    LireEmissionsPrestations ws, idSaisie, wsEpB
    LireChiffresCles ws, idSaisie, wsCcB

    RafraichirCopies
    MasquerBrutes
    ws.Activate
    Application.ScreenUpdating = True

    If MsgBox(societe & " (" & anN & ") enregistree." & vbLf & vbLf & _
              "Vider la feuille Collecte pour la saisie suivante ?", vbYesNo + vbInformation, "Collecte") = vbYes Then
        ws.Range("A1:F66").ClearContents
    End If
    Exit Sub

Erreur:
    numErr = Err.Number
    descErr = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Erreur " & numErr & " :" & vbLf & descErr, vbCritical, "Collecte"
End Sub

' ---------------------------------------------------------------------------------
' Controles de la feuille Collecte (A1:F66). Renvoie "" si tout est correct.
' ---------------------------------------------------------------------------------
Private Function ControlerCollecte(ByVal ws As Worksheet) As String
    Dim msg As String, c As Range, invalides As String, v As Variant

    ' Structure : libelles fixes du modele
    If UCase$(Texte(ws.Range("A1").Value)) <> "NOM DE LA SOCIETE" _
       Or InStr(1, Texte(ws.Range("B9").Value), "EMISSIONS", vbTextCompare) = 0 _
       Or InStr(1, Texte(ws.Range("B29").Value), "PRESTATIONS", vbTextCompare) = 0 _
       Or InStr(1, Texte(ws.Range("B49").Value), "AUTRES CHIFFRES", vbTextCompare) = 0 _
       Or UCase$(Texte(ws.Range("A10").Value)) <> "BRANCHES" _
       Or UCase$(Texte(ws.Range("A50").Value)) <> "RUBRIQUES" Then
        ControlerCollecte = "- Le bloc colle en A1 n'a pas la mise en page du modele (A1:F66)."
        Exit Function
    End If

    ' Champs obligatoires
    If Len(Texte(ws.Range("B1").Value)) = 0 Then msg = msg & vbLf & "- Nom de la societe vide (B1)."
    If Len(Texte(ws.Range("B2").Value)) = 0 Then msg = msg & vbLf & "- Pays vide (B2)."
    v = ws.Range("B10").Value
    If IsError(v) Then
        msg = msg & vbLf & "- Annee N-1 invalide (B10)."
    ElseIf Not IsNumeric(v) Or IsEmpty(v) Then
        msg = msg & vbLf & "- Annee N-1 invalide (B10)."
    ElseIf v < 1990 Or v > 2100 Or v <> Int(v) Then
        msg = msg & vbLf & "- Annee N-1 invalide (B10)."
    End If

    ' Date de creation : vide ou date
    v = ws.Range("B4").Value
    If IsError(v) Then
        msg = msg & vbLf & "- Date de creation invalide (B4)."
    ElseIf Not IsEmpty(v) And VarType(v) <> vbDate And Not IsNumeric(v) And Not IsDate(v) Then
        msg = msg & vbLf & "- Date de creation invalide (B4)."
    End If

    ' Cellules chiffrees : vides ou numeriques
    For Each c In ws.Range(CELLULES_NUMERIQUES).Cells
        v = c.Value
        If IsError(v) Then
            invalides = invalides & " " & c.Address(False, False)
        ElseIf Not IsEmpty(v) And Not IsNumeric(v) Then
            invalides = invalides & " " & c.Address(False, False)
        End If
    Next c
    If Len(invalides) > 0 Then msg = msg & vbLf & "- Valeurs non numeriques :" & invalides

    If Len(msg) > 0 Then msg = Mid$(msg, 2)
    ControlerCollecte = msg
End Function

' ---------------------------------------------------------------------------------
' Lecture du formulaire vers les feuilles brutes
' ---------------------------------------------------------------------------------
Private Function EntetesIdentification() As Variant
    EntetesIdentification = Array("Id_saisie", "Cle", "Societe", "Pays", "DG", "Date_creation", _
        "Capital_social", "Cadres", "Maitrise", "Employes", "Annee_N", "Date_import")
End Function

Private Function EntetesEmissionsPrestations() As Variant
    EntetesEmissionsPrestations = Array("Id_saisie", "Cle", "Societe", "Pays", "Bloc", "Categorie", _
        "Rubrique", "Annee", "Valeur", "Cellule_source")
End Function

Private Function EntetesChiffresCles() As Variant
    EntetesChiffresCles = Array("Id_saisie", "Cle", "Societe", "Pays", "Rubrique", "Annee", _
        "Valeur", "Cellule_source")
End Function

' Une ligne par saisie.
Private Sub LireIdentification(ByVal ws As Worksheet, ByVal idSaisie As Long, ByVal dest As Worksheet)
    AjouterLigne dest, Array(idSaisie, _
        Cle(Texte(ws.Range("B2").Value), Texte(ws.Range("B1").Value), CLng(ws.Range("B10").Value) + 1), _
        Texte(ws.Range("B1").Value), _
        Texte(ws.Range("B2").Value), _
        Texte(ws.Range("B3").Value), _
        ValeurDate(ws.Range("B4").Value), _
        ValeurCapital(ws.Range("B5").Value), _
        Nombre(ws.Range("D6").Value), _
        Nombre(ws.Range("D7").Value), _
        Nombre(ws.Range("D8").Value), _
        CLng(ws.Range("B10").Value) + 1, _
        Now)
End Sub

' Blocs A (emissions) et B (prestations, meme grille 20 lignes plus bas) :
' N-1 en colonne B, N en colonne D. Une ligne par rubrique et par annee.
Private Sub LireEmissionsPrestations(ByVal ws As Worksheet, ByVal idSaisie As Long, ByVal dest As Worksheet)
    Dim lignes As Variant, blocs As Variant
    Dim i As Long, b As Long, r As Long, anN1 As Long
    Dim societe As String, pays As String, categorie As String, rubrique As String

    societe = Texte(ws.Range("B1").Value)
    pays = Texte(ws.Range("B2").Value)
    anN1 = CLng(ws.Range("B10").Value)

    lignes = Array(13, 14, 15, 16, 17, 18, 20, 21, 22, 23, 24, 25, 27)
    blocs = Array("Emissions", "Prestations")
    For b = 0 To 1
        For i = LBound(lignes) To UBound(lignes)
            r = lignes(i) + 20 * b
            Select Case lignes(i)
                Case Is < 19: categorie = "Individuelles"
                Case Is < 26: categorie = "Collectives"
                Case Else: categorie = "Acceptations"
            End Select
            rubrique = Libelle(ws.Cells(r, 1).Value)
            AjouterLigne dest, Array(idSaisie, Cle(pays, societe, blocs(b), categorie, rubrique, anN1), _
                societe, pays, blocs(b), categorie, rubrique, anN1, Nombre(ws.Cells(r, 2).Value), ws.Cells(r, 2).Address(False, False))
            AjouterLigne dest, Array(idSaisie, Cle(pays, societe, blocs(b), categorie, rubrique, anN1 + 1), _
                societe, pays, blocs(b), categorie, rubrique, anN1 + 1, Nombre(ws.Cells(r, 4).Value), ws.Cells(r, 4).Address(False, False))
        Next i
    Next b
End Sub

' Bloc C : N-1 en colonne C, N en colonne D, lignes 51 a 64.
' Les lignes 65 (autres actifs) et 66 (taux de couverture) sont des formules, non reprises.
Private Sub LireChiffresCles(ByVal ws As Worksheet, ByVal idSaisie As Long, ByVal dest As Worksheet)
    Dim r As Long, anN1 As Long
    Dim societe As String, pays As String, rubrique As String

    societe = Texte(ws.Range("B1").Value)
    pays = Texte(ws.Range("B2").Value)
    anN1 = CLng(ws.Range("B10").Value)

    For r = 51 To 64
        rubrique = Libelle(ws.Cells(r, 1).Value)
        AjouterLigne dest, Array(idSaisie, Cle(pays, societe, rubrique, anN1), _
            societe, pays, rubrique, anN1, Nombre(ws.Cells(r, 3).Value), ws.Cells(r, 3).Address(False, False))
        AjouterLigne dest, Array(idSaisie, Cle(pays, societe, rubrique, anN1 + 1), _
            societe, pays, rubrique, anN1 + 1, Nombre(ws.Cells(r, 4).Value), ws.Cells(r, 4).Address(False, False))
    Next r
End Sub

Private Sub AjouterLigne(ByVal dest As Worksheet, ByVal valeurs As Variant)
    Dim lig As Long
    lig = dest.Cells(dest.Rows.Count, 1).End(xlUp).Row + 1
    dest.Cells(lig, 1).Resize(1, UBound(valeurs) - LBound(valeurs) + 1).Value = valeurs
End Sub

' ---------------------------------------------------------------------------------
' Doublons : meme societe (sans tenir compte de la casse) et meme annee N
' ---------------------------------------------------------------------------------
Private Function SaisiesExistantes(ByVal wsIdB As Worksheet, ByVal societe As String, ByVal anN As Long) As Collection
    Dim res As New Collection, r As Long, der As Long

    der = wsIdB.Cells(wsIdB.Rows.Count, 1).End(xlUp).Row
    For r = 2 To der
        If StrComp(Texte(wsIdB.Cells(r, 3).Value), societe, vbTextCompare) = 0 _
           And Val(wsIdB.Cells(r, 11).Value) = anN Then
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

    wsIdB.Columns("F").NumberFormat = "dd/mm/yyyy"
    wsIdB.Columns("G").NumberFormat = "#,##0"
    wsIdB.Columns("L").NumberFormat = "dd/mm/yyyy hh:mm"
    wsEpB.Columns("I").NumberFormat = "#,##0.000"
    wsCcB.Columns("G").NumberFormat = "#,##0.000"

    ' On saute la colonne A (Id_saisie) et on s'arrete avant Date_import / Cellule_source ;
    ' la valeur est reprise par formule * 1000.
    CreerCopie wsIdB, F_ID, Array("Cle", "Societe", "Pays", "DG", "Date_creation", "Capital_social", _
        "Cadres", "Maitrise", "Employes", "Annee_N"), 10, ""
    CreerCopie wsEpB, F_EP, Array("Cle", "Societe", "Pays", "Bloc", "Categorie", "Rubrique", "Annee", _
        "Valeur_FCFA"), 7, "I"
    CreerCopie wsCcB, NomChiffresCles(), Array("Cle", "Societe", "Pays", "Rubrique", "Annee", _
        "Valeur_FCFA"), 5, "G"
End Sub

' Copie visible d'une feuille brute : colonnes B a (1 + nbCols) en valeurs, puis,
' si colValeur est renseignee, une colonne Valeur_FCFA = brute!colValeur * 1000.
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
        ws.Columns("E").NumberFormat = "dd/mm/yyyy"
        ws.Columns("F").NumberFormat = "#,##0"
    End If
    ws.Visible = xlSheetVisible
    ws.Range("A1").Resize(1, nbEntetes).EntireColumn.AutoFit
End Sub

' ---------------------------------------------------------------------------------
' Conversions
' ---------------------------------------------------------------------------------
' Cle de recherche : elements separes par "|", ex. "BENIN|ATLANTIQUE ASSURANCE BENIN VIE|2021".
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

' Libelle de rubrique nettoye (espaces, fautes de frappe du modele, majuscule initiale).
Private Function Libelle(ByVal v As Variant) As String
    Dim s As String
    s = Texte(v)
    s = Replace(s, "autes actifs", "autres actifs", , , vbTextCompare)
    s = Replace(s, "propres er ", "propres et ", , , vbTextCompare)
    If Len(s) > 0 Then s = UCase$(Left$(s, 1)) & Mid$(s, 2)
    Libelle = s
End Function

' Cellule vide ou en erreur -> Empty (distinct d'un vrai 0).
Private Function Nombre(ByVal v As Variant) As Variant
    If IsError(v) Then
        Nombre = Empty
    ElseIf IsEmpty(v) Then
        Nombre = Empty
    ElseIf Len(Trim$(CStr(v))) = 0 Then
        Nombre = Empty
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
