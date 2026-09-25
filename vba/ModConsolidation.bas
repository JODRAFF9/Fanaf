Attribute VB_Name = "ModConsolidation"
Option Explicit

' Consolidation des classeurs au format "Modele.xlsx" (formulaire sur la 1re feuille,
' montants en milliers de F CFA).
'
' Feuilles brutes (masquees), donnees telles que saisies :
'   Identification_brut, Emission&Prestations_brut, Chiffres cles_brut
' Copies visibles, sans Fichier, Date_import ni Cellule_source :
'   Identification, Emission&Prestations, Chiffres cles
'   Valeur_FCFA = valeur brute (milliers) * 1000, par formule
'
' Le code source est sans accents pour s'importer sans souci dans l'editeur VBA ;
' le "e accent aigu" du nom de feuille est produit par ChrW(233).

Private Const F_ID As String = "Identification"
Private Const F_EP As String = "Emission&Prestations"
Private Const SUFFIXE_BRUT As String = "_brut"

Private Function NomChiffresCles() As String
    NomChiffresCles = "Chiffres cl" & ChrW(233) & "s"
End Function

Public Sub ConsoliderClasseurs()
    Dim dossier As String, fichier As String, rejets As String, msg As String
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wsIdB As Worksheet, wsEpB As Worksheet, wsCcB As Worksheet
    Dim ligId As Long, ligEp As Long, ligCc As Long, nbOk As Long
    Dim calcInitial As XlCalculation
    Dim numErr As Long, descErr As String

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choisir le dossier des classeurs a consolider"
        If .Show <> -1 Then Exit Sub
        dossier = .SelectedItems(1)
    End With
    If Right$(dossier, 1) <> Application.PathSeparator Then dossier = dossier & Application.PathSeparator

    calcInitial = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    On Error GoTo Erreur

    Set wsIdB = PreparerFeuille(F_ID & SUFFIXE_BRUT, Array("Fichier", "Societe", "Pays", "DG", _
        "Date_creation", "Capital_social", "Cadres", "Maitrise", "Employes", "Annee_N", "Date_import"))
    Set wsEpB = PreparerFeuille(F_EP & SUFFIXE_BRUT, Array("Fichier", "Societe", "Pays", "Bloc", _
        "Categorie", "Rubrique", "Annee", "Valeur", "Cellule_source"))
    Set wsCcB = PreparerFeuille(NomChiffresCles() & SUFFIXE_BRUT, Array("Fichier", "Societe", "Pays", _
        "Rubrique", "Annee", "Valeur", "Cellule_source"))
    ligId = 2
    ligEp = 2
    ligCc = 2

    fichier = Dir(dossier & "*.xls*")
    Do While Len(fichier) > 0
        If fichier <> ThisWorkbook.Name And Left$(fichier, 2) <> "~$" Then
            Set wbSrc = Workbooks.Open(Filename:=dossier & fichier, UpdateLinks:=0, ReadOnly:=True)
            Set wsSrc = wbSrc.Worksheets(1)
            If FormatValide(wsSrc) Then
                LireIdentification wsSrc, fichier, wsIdB, ligId
                LireEmissionsPrestations wsSrc, fichier, wsEpB, ligEp
                LireChiffresCles wsSrc, fichier, wsCcB, ligCc
                nbOk = nbOk + 1
            Else
                rejets = rejets & vbLf & "  " & fichier
            End If
            wbSrc.Close SaveChanges:=False
            Set wbSrc = Nothing
        End If
        fichier = Dir()
    Loop

    ' Copies visibles : on saute la colonne A (Fichier) et on s'arrete avant
    ' Date_import / Cellule_source ; la valeur est reprise par formule * 1000.
    CreerCopie wsIdB, F_ID, Array("Societe", "Pays", "DG", "Date_creation", "Capital_social", _
        "Cadres", "Maitrise", "Employes", "Annee_N"), 9, "", ligId - 1
    CreerCopie wsEpB, F_EP, Array("Societe", "Pays", "Bloc", "Categorie", "Rubrique", "Annee", _
        "Valeur_FCFA"), 6, "H", ligEp - 1
    CreerCopie wsCcB, NomChiffresCles(), Array("Societe", "Pays", "Rubrique", "Annee", _
        "Valeur_FCFA"), 4, "F", ligCc - 1

    MettreEnForme wsIdB, wsEpB, wsCcB
    wsIdB.Visible = xlSheetHidden
    wsEpB.Visible = xlSheetHidden
    wsCcB.Visible = xlSheetHidden

    Application.Calculation = calcInitial
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    msg = nbOk & " classeur(s) consolide(s)."
    If Len(rejets) > 0 Then msg = msg & vbLf & vbLf & "Ignores (format non reconnu) :" & rejets
    MsgBox msg, vbInformation, "Consolidation"
    Exit Sub

Erreur:
    numErr = Err.Number
    descErr = Err.Description
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.Calculation = calcInitial
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Erreur " & numErr & " sur le fichier " & fichier & " :" & vbLf & descErr, vbCritical, "Consolidation"
End Sub

' Verifie que la feuille a bien la mise en page du modele.
Private Function FormatValide(ByVal ws As Worksheet) As Boolean
    FormatValide = UCase$(Texte(ws.Range("A1").Value)) = "NOM DE LA SOCIETE" _
        And InStr(1, Texte(ws.Range("B9").Value), "EMISSIONS", vbTextCompare) > 0 _
        And InStr(1, Texte(ws.Range("B29").Value), "PRESTATIONS", vbTextCompare) > 0 _
        And InStr(1, Texte(ws.Range("B49").Value), "AUTRES CHIFFRES", vbTextCompare) > 0 _
        And IsNumeric(ws.Range("B10").Value) And Not IsEmpty(ws.Range("B10").Value)
End Function

' Une ligne par classeur.
Private Sub LireIdentification(ByVal ws As Worksheet, ByVal fichier As String, ByVal dest As Worksheet, ByRef lig As Long)
    EcrireLigne dest, lig, Array(fichier, _
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
Private Sub LireEmissionsPrestations(ByVal ws As Worksheet, ByVal fichier As String, ByVal dest As Worksheet, ByRef lig As Long)
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
            EcrireLigne dest, lig, Array(fichier, societe, pays, blocs(b), categorie, rubrique, _
                anN1, Nombre(ws.Cells(r, 2).Value), ws.Cells(r, 2).Address(False, False))
            EcrireLigne dest, lig, Array(fichier, societe, pays, blocs(b), categorie, rubrique, _
                anN1 + 1, Nombre(ws.Cells(r, 4).Value), ws.Cells(r, 4).Address(False, False))
        Next i
    Next b
End Sub

' Bloc C : N-1 en colonne C, N en colonne D, lignes 51 a 64.
' Les lignes 65 (autres actifs) et 66 (taux de couverture) sont des formules, non reprises.
Private Sub LireChiffresCles(ByVal ws As Worksheet, ByVal fichier As String, ByVal dest As Worksheet, ByRef lig As Long)
    Dim r As Long, anN1 As Long
    Dim societe As String, pays As String, rubrique As String

    societe = Texte(ws.Range("B1").Value)
    pays = Texte(ws.Range("B2").Value)
    anN1 = CLng(ws.Range("B10").Value)

    For r = 51 To 64
        rubrique = Libelle(ws.Cells(r, 1).Value)
        EcrireLigne dest, lig, Array(fichier, societe, pays, rubrique, _
            anN1, Nombre(ws.Cells(r, 3).Value), ws.Cells(r, 3).Address(False, False))
        EcrireLigne dest, lig, Array(fichier, societe, pays, rubrique, _
            anN1 + 1, Nombre(ws.Cells(r, 4).Value), ws.Cells(r, 4).Address(False, False))
    Next r
End Sub

Private Sub EcrireLigne(ByVal dest As Worksheet, ByRef lig As Long, ByVal valeurs As Variant)
    dest.Cells(lig, 1).Resize(1, UBound(valeurs) - LBound(valeurs) + 1).Value = valeurs
    lig = lig + 1
End Sub

' Copie visible d'une feuille brute : colonnes B a (1 + nbCols) en valeurs, puis,
' si colValeur est renseignee, une colonne Valeur_FCFA = brute!colValeur * 1000.
Private Sub CreerCopie(ByVal wsBrut As Worksheet, ByVal nom As String, ByVal entetes As Variant, _
                       ByVal nbCols As Long, ByVal colValeur As String, ByVal derLig As Long)
    Dim ws As Worksheet, nbEntetes As Long

    Set ws = PreparerFeuille(nom, entetes)
    nbEntetes = UBound(entetes) - LBound(entetes) + 1

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
        ws.Columns("D").NumberFormat = "dd/mm/yyyy"
        ws.Columns("E").NumberFormat = "#,##0"
    End If
    ws.Visible = xlSheetVisible
    ws.Range("A1").Resize(1, nbEntetes).EntireColumn.AutoFit
End Sub

Private Sub MettreEnForme(ByVal wsIdB As Worksheet, ByVal wsEpB As Worksheet, ByVal wsCcB As Worksheet)
    wsIdB.Columns("E").NumberFormat = "dd/mm/yyyy"
    wsIdB.Columns("F").NumberFormat = "#,##0"
    wsIdB.Columns("K").NumberFormat = "dd/mm/yyyy hh:mm"
    wsEpB.Columns("H").NumberFormat = "#,##0.000"
    wsCcB.Columns("F").NumberFormat = "#,##0.000"
    wsIdB.Columns("A:K").AutoFit
    wsEpB.Columns("A:I").AutoFit
    wsCcB.Columns("A:G").AutoFit
End Sub

' Renvoie la feuille demandee (creee si absente), videe, avec sa ligne d'en-tetes.
Private Function PreparerFeuille(ByVal nom As String, ByVal entetes As Variant) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nom
    End If

    ws.Cells.Clear
    With ws.Range("A1").Resize(1, UBound(entetes) - LBound(entetes) + 1)
        .Value = entetes
        .Font.Bold = True
    End With
    Set PreparerFeuille = ws
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
