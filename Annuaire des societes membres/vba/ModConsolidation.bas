Attribute VB_Name = "ModConsolidation"
Option Explicit

' Collecte des formulaires FANAF, vie et non-vie (montants en milliers de F CFA).
'
' Utilisation :
'   1. Lancer une fois InitialiserClasseur : cree les feuilles Collecte Vie et
'      Collecte Non Vie, chacune avec ses boutons Enregistrer (H2) et
'      Supprimer le precedent (H5).
'   2. Copier le formulaire (colonnes A a F) et le coller en A1 de la feuille de collecte
'      correspondant a la branche de la societe.
'   3. Cliquer sur Enregistrer : la saisie est controlee puis ajoutee a la base.
'      La branche (Vie / Non-vie) est celle de la feuille de collecte.
'      Chaque enregistrement recoit un numero qui augmente a chaque fois (jamais reutilise).
'   4. Supprimer le precedent : efface le dernier enregistrement de la branche de la feuille.
'
' Feuilles brutes (masquees), donnees telles que saisies :
'   Identification (brut), Emission&Prestations (brut), Chiffres cles (brut)
' Copies visibles, sans Date d'import ni Cellule source :
'   Identification, Emission&Prestations, Chiffres cles
'   Valeur (F CFA) = valeur brute (milliers F CFA) * 1000, par formule
'
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
' Code source en ASCII pur : aucun probleme d'encodage a l'import ou au copier-coller.
' Le "e accent aigu" du nom de feuille Chiffres cles est produit par ChrW(233).

Private Const F_COLLECTE_VIE As String = "Collecte Vie"
Private Const F_COLLECTE_NV As String = "Collecte Non Vie"
Private Const TYPE_VIE As String = "Vie"
Private Const TYPE_NV As String = "Non-vie"
Private Const F_ID As String = "Identification"
Private Const F_EP As String = "Emission&Prestations"
Private Const SUFFIXE_BRUT As String = " (brut)"
Private Const NOM_BOUTON As String = "btnEnregistrer"
Private Const NOM_BOUTON_SUPPR As String = "btnSupprimer"
Private Const NOM_COMPTEUR As String = "NumeroEnregistrement"   ' nom masque du classeur

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
Private mType As String ' type de la feuille de collecte en cours : Vie ou Non-vie

Private Function NomChiffresCles() As String
    NomChiffresCles = "Chiffres cl" & ChrW(233) & "s"
End Function

' ---------------------------------------------------------------------------------
' Mise en place : feuille Collecte + bouton Enregistrer
' ---------------------------------------------------------------------------------
Public Sub InitialiserClasseur()
    PreparerCollecte F_COLLECTE_NV, "EnregistrerNonVie", "SupprimerPrecedentNonVie"
    PreparerCollecte F_COLLECTE_VIE, "EnregistrerVie", "SupprimerPrecedentVie"

    FeuilleBrute F_ID & SUFFIXE_BRUT, EntetesIdentification()
    FeuilleBrute F_EP & SUFFIXE_BRUT, EntetesEmissionsPrestations()
    FeuilleBrute NomChiffresCles() & SUFFIXE_BRUT, EntetesChiffresCles()
    RafraichirCopies
    MasquerBrutes

    ThisWorkbook.Worksheets(F_COLLECTE_VIE).Activate
    MsgBox "Feuilles " & F_COLLECTE_VIE & " et " & F_COLLECTE_NV & " pretes : collez le formulaire en A1 " & _
        "puis cliquez sur Enregistrer.", vbInformation, "Collecte"
End Sub

' Cree (si besoin) une feuille de collecte avec ses boutons Enregistrer et Supprimer le precedent.
Private Sub PreparerCollecte(ByVal nom As String, ByVal macroEnregistrer As String, ByVal macroSupprimer As String)
    Dim ws As Worksheet, cible As Range, btn As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = nom
    End If

    On Error Resume Next
    ws.Buttons(NOM_BOUTON).Delete
    ws.Buttons(NOM_BOUTON_SUPPR).Delete
    On Error GoTo 0

    ' Nettoie la cellule Code societe des versions precedentes
    ws.Range("H4:H5").Clear

    Set cible = ws.Range("H2")
    Set btn = ws.Buttons.Add(cible.Left, cible.Top, 150, 28)
    btn.Name = NOM_BOUTON
    btn.Caption = "Enregistrer"
    btn.OnAction = macroEnregistrer

    Set cible = ws.Range("H5")
    Set btn = ws.Buttons.Add(cible.Left, cible.Top, 150, 28)
    btn.Name = NOM_BOUTON_SUPPR
    btn.Caption = "Supprimer le pr" & ChrW(233) & "c" & ChrW(233) & "dent"
    btn.OnAction = macroSupprimer
End Sub

' ---------------------------------------------------------------------------------
' Bouton Enregistrer
' ---------------------------------------------------------------------------------
Public Sub EnregistrerVie()
    EnregistrerCollecte ThisWorkbook.Worksheets(F_COLLECTE_VIE), TYPE_VIE
End Sub

Public Sub EnregistrerNonVie()
    EnregistrerCollecte ThisWorkbook.Worksheets(F_COLLECTE_NV), TYPE_NV
End Sub

Private Sub EnregistrerCollecte(ByVal ws As Worksheet, ByVal typeCollecte As String)
    Dim wsIdB As Worksheet, wsEpB As Worksheet, wsCcB As Worksheet
    Dim erreurs As String, societe As String, pays As String, idSaisie As Long, anN As Long
    Dim anciens As Collection, entrees As Collection
    Dim numErr As Long, descErr As String

    mType = typeCollecte
    erreurs = ControlerCollecte(ws)
    If Len(erreurs) > 0 Then
        MsgBox "Enregistrement refuse :" & vbLf & erreurs, vbExclamation, "Collecte"
        Exit Sub
    End If

    societe = Texte(ws.Range("B1").Value)
    pays = Texte(ws.Range("B2").Value)
    anN = AnneeN1(ws) + 1

    Set wsIdB = FeuilleBrute(F_ID & SUFFIXE_BRUT, EntetesIdentification())
    Set wsEpB = FeuilleBrute(F_EP & SUFFIXE_BRUT, EntetesEmissionsPrestations())
    Set wsCcB = FeuilleBrute(NomChiffresCles() & SUFFIXE_BRUT, EntetesChiffresCles())

    ' Meme societe, meme pays et meme annee deja enregistres : remplacement sur confirmation
    Set anciens = SaisiesExistantes(wsIdB, societe, pays, anN)
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

    idSaisie = ProchainNumero(wsIdB)
    Set entrees = LireEntrees(ws)
    EcrireIdentification ws, idSaisie, wsIdB
    EcrireEntrees ws, idSaisie, entrees, wsEpB, wsCcB

    RafraichirCopies
    MasquerBrutes
    ws.Activate
    Application.ScreenUpdating = True

    If MsgBox(societe & " (" & anN & ") enregistree sous le numero " & idSaisie & "." & vbLf & vbLf & _
              "Vider la feuille " & ws.Name & " pour la saisie suivante ?", vbYesNo + vbInformation, "Collecte") = vbYes Then
        ws.Range("A1:F" & LIGNE_MAX).ClearContents
    End If
    Exit Sub

Erreur:
    numErr = Err.Number
    descErr = Err.Description
    Application.ScreenUpdating = True
    MsgBox "Erreur " & numErr & " :" & vbLf & descErr, vbCritical, "Collecte"
End Sub

' ---------------------------------------------------------------------------------
' Numero d'enregistrement : compteur conserve dans un nom masque du classeur,
' pour que le numero augmente toujours, meme apres une suppression.
' ---------------------------------------------------------------------------------
Private Function ProchainNumero(ByVal wsIdB As Worksheet) As Long
    Dim n As Long, v As Variant, maxi As Double

    On Error Resume Next
    v = Evaluate(ThisWorkbook.Names(NOM_COMPTEUR).RefersTo)
    On Error GoTo 0
    If Not IsError(v) Then
        If IsNumeric(v) And Not IsEmpty(v) Then n = CLng(v)
    End If

    maxi = Application.WorksheetFunction.Max(wsIdB.Columns(1))
    If n < maxi Then n = CLng(maxi)

    n = n + 1
    ThisWorkbook.Names.Add Name:=NOM_COMPTEUR, RefersTo:="=" & n, Visible:=False
    ProchainNumero = n
End Function

' ---------------------------------------------------------------------------------
' Bouton Supprimer le precedent : dernier enregistrement de la branche de la feuille
' ---------------------------------------------------------------------------------
Public Sub SupprimerPrecedentVie()
    SupprimerPrecedent ThisWorkbook.Worksheets(F_COLLECTE_VIE), TYPE_VIE
End Sub

Public Sub SupprimerPrecedentNonVie()
    SupprimerPrecedent ThisWorkbook.Worksheets(F_COLLECTE_NV), TYPE_NV
End Sub

Private Sub SupprimerPrecedent(ByVal ws As Worksheet, ByVal typeCollecte As String)
    Dim wsIdB As Worksheet, wsEpB As Worksheet, wsCcB As Worksheet
    Dim r As Long, der As Long, ligne As Long, numero As Double
    Dim ids As New Collection

    Set wsIdB = FeuilleBrute(F_ID & SUFFIXE_BRUT, EntetesIdentification())
    Set wsEpB = FeuilleBrute(F_EP & SUFFIXE_BRUT, EntetesEmissionsPrestations())
    Set wsCcB = FeuilleBrute(NomChiffresCles() & SUFFIXE_BRUT, EntetesChiffresCles())

    ' Plus grand numero de la branche
    der = wsIdB.Cells(wsIdB.Rows.Count, 1).End(xlUp).Row
    For r = 2 To der
        If StrComp(Texte(wsIdB.Cells(r, 4).Value), typeCollecte, vbTextCompare) = 0 _
           And IsNumeric(wsIdB.Cells(r, 1).Value) Then
            If wsIdB.Cells(r, 1).Value > numero Then
                numero = wsIdB.Cells(r, 1).Value
                ligne = r
            End If
        End If
    Next r

    If ligne = 0 Then
        MsgBox "Aucun enregistrement " & typeCollecte & " a supprimer.", vbInformation, "Collecte"
        Exit Sub
    End If

    If MsgBox("Supprimer l'enregistrement numero " & numero & " :" & vbLf & _
              Texte(wsIdB.Cells(ligne, 2).Value) & " (" & Texte(wsIdB.Cells(ligne, 3).Value) & ", " & _
              wsIdB.Cells(ligne, 11).Value & ") ?", vbYesNo + vbExclamation, "Collecte") = vbNo Then Exit Sub

    Application.ScreenUpdating = False
    ids.Add numero
    SupprimerSaisies wsIdB, ids
    SupprimerSaisies wsEpB, ids
    SupprimerSaisies wsCcB, ids
    RafraichirCopies
    MasquerBrutes
    ws.Activate
    Application.ScreenUpdating = True

    MsgBox "Enregistrement numero " & numero & " supprime.", vbInformation, "Collecte"
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

' Type du formulaire d'apres le titre du bloc B : SINISTRALITE -> Non-vie, sinon Vie.
Private Function TypeFormulaire(ByVal ws As Worksheet) As String
    If InStr(1, Texte(ws.Cells(mB, 2).Value), "SINISTRALITE", vbTextCompare) > 0 Then
        TypeFormulaire = TYPE_NV
    Else
        TypeFormulaire = TYPE_VIE
    End If
End Function

Private Function NomBloc(ByVal titre As String) As String
    If InStr(1, titre, "EMISSIONS", vbTextCompare) > 0 Then
        NomBloc = "Emissions"
    ElseIf InStr(1, titre, "PRESTATIONS", vbTextCompare) > 0 Then
        NomBloc = "Prestations"
    ElseIf InStr(1, titre, "SINISTRALITE", vbTextCompare) > 0 Then
        NomBloc = "Sinistralite"
    Else
        NomBloc = titre
    End If
End Function

' Mesure unique d'un bloc vie, d'apres son titre.
Private Function MesureParDefaut(ByVal titre As String) As String
    If InStr(1, titre, "EMISSIONS", vbTextCompare) > 0 Then
        MesureParDefaut = "Emissions nettes"
    ElseIf InStr(1, titre, "PRESTATIONS", vbTextCompare) > 0 Then
        MesureParDefaut = "Prestations versees"
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
        ControlerCollecte = "- Le bloc colle en A1 n'a pas la mise en page du formulaire" & vbLf & _
            "  (titres A - EMISSIONS, B - PRESTATIONS ou SINISTRALITE et C - AUTRES CHIFFRES" & vbLf & _
            "  introuvables, ou lignes Branches / Rubriques absentes)."
        Exit Function
    End If

    ' Le formulaire colle doit correspondre a la feuille de collecte
    If TypeFormulaire(ws) <> mType Then
        ControlerCollecte = "- Ce formulaire est un formulaire " & TypeFormulaire(ws) & _
            " : collez-le dans la feuille " & IIf(TypeFormulaire(ws) = TYPE_VIE, F_COLLECTE_VIE, F_COLLECTE_NV) & "."
        Exit Function
    End If

    ' Champs obligatoires
    If Len(Texte(ws.Range("B1").Value)) = 0 Then msg = msg & vbLf & "- Nom de la societe vide (B1)."
    If Len(Texte(ws.Range("B2").Value)) = 0 Then msg = msg & vbLf & "- Pays vide (B2)."
    v = Nombre(ws.Cells(mA + 1, 2).Value)
    If Not IsNumeric(v) Or IsEmpty(v) Then
        msg = msg & vbLf & "- Annee N-1 invalide (" & ws.Cells(mA + 1, 2).Address(False, False) & ")."
        ControlerCollecte = Mid$(msg, 2)
        Exit Function
    ElseIf v < 1990 Or v > 2100 Or v <> Int(v) Then
        msg = msg & vbLf & "- Annee N-1 invalide (" & ws.Cells(mA + 1, 2).Address(False, False) & ")."
        ControlerCollecte = Mid$(msg, 2)
        Exit Function
    End If

    ' Date de creation : vide ou date
    v = ws.Range("B4").Value
    If IsError(v) Then
        msg = msg & vbLf & "- Date de creation invalide (B4)."
    ElseIf Not IsEmpty(v) And VarType(v) <> vbDate And Not IsNumeric(v) And Not IsDate(v) Then
        msg = msg & vbLf & "- Date de creation invalide (B4)."
    End If

    ' Effectifs : vides ou numeriques
    For Each c In ws.Range("D6:D8").Cells
        If Not EstNombreOuVide(c.Value) Then invalides = invalides & " " & c.Address(False, False)
    Next c

    ' Cellules chiffrees des blocs : vides ou numeriques (nombres colles en texte acceptes)
    Set entrees = LireEntrees(ws)
    If entrees.Count = 0 Then msg = msg & vbLf & "- Aucune donnee trouvee dans les blocs A, B et C."
    For Each e In entrees
        Set c = e(6)
        If Not EstNombreOuVide(c.Value) Then invalides = invalides & " " & c.Address(False, False)
    Next e
    If Len(invalides) > 0 Then msg = msg & vbLf & "- Valeurs non numeriques :" & invalides

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
    EntetesIdentification = Array("No enregistrement", "Societe", "Pays", "Branche", _
        "Directeur general", "Date de creation", "Capital social (F CFA)", "Cadres", "Maitrise", _
        "Employes", "Annee N", "Date d'import")
End Function

Private Function EntetesEmissionsPrestations() As Variant
    EntetesEmissionsPrestations = Array("No enregistrement", "Societe", "Pays", "Branche", _
        "Bloc", "Categorie", "Rubrique", "Mesure", "Annee", "Valeur (milliers F CFA)", "Cellule source")
End Function

Private Function EntetesChiffresCles() As Variant
    EntetesChiffresCles = Array("No enregistrement", "Societe", "Pays", "Branche", _
        "Rubrique", "Annee", "Valeur (milliers F CFA)", "Cellule source")
End Function

' Une ligne par saisie.
Private Sub EcrireIdentification(ByVal ws As Worksheet, ByVal idSaisie As Long, ByVal dest As Worksheet)
    AjouterLigne dest, Array(idSaisie, _
        Texte(ws.Range("B1").Value), _
        Texte(ws.Range("B2").Value), _
        mType, _
        Texte(ws.Range("B3").Value), _
        ValeurDate(ws.Range("B4").Value), _
        ValeurCapital(ws.Range("B5").Value), _
        Nombre(ws.Range("D6").Value), _
        Nombre(ws.Range("D7").Value), _
        Nombre(ws.Range("D8").Value), _
        AnneeN1(ws) + 1, _
        Now)
End Sub

' Une ligne par entree, dans Emission&Prestations ou Chiffres cles.
Private Sub EcrireEntrees(ByVal ws As Worksheet, ByVal idSaisie As Long, ByVal entrees As Collection, _
                          ByVal wsEpB As Worksheet, ByVal wsCcB As Worksheet)
    Dim e As Variant, c As Range
    Dim societe As String, pays As String, typ As String

    societe = Texte(ws.Range("B1").Value)
    pays = Texte(ws.Range("B2").Value)
    typ = mType

    For Each e In entrees
        Set c = e(6)
        If e(0) = T_EP Then
            AjouterLigne wsEpB, Array(idSaisie, _
                societe, pays, typ, e(1), e(2), e(3), e(4), e(5), _
                Nombre(c.Value), c.Address(False, False))
        Else
            AjouterLigne wsCcB, Array(idSaisie, _
                societe, pays, typ, e(3), e(5), _
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
' Doublons : meme societe, meme pays (sans tenir compte de la casse) et meme annee N
' ---------------------------------------------------------------------------------
Private Function SaisiesExistantes(ByVal wsIdB As Worksheet, ByVal societe As String, ByVal pays As String, _
                                   ByVal anN As Long) As Collection
    Dim res As New Collection, r As Long, der As Long

    der = wsIdB.Cells(wsIdB.Rows.Count, 1).End(xlUp).Row
    For r = 2 To der
        If StrComp(Texte(wsIdB.Cells(r, 2).Value), societe, vbTextCompare) = 0 _
           And StrComp(Texte(wsIdB.Cells(r, 3).Value), pays, vbTextCompare) = 0 _
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
    wsEpB.Columns("J").NumberFormat = "#,##0.000"
    wsCcB.Columns("G").NumberFormat = "#,##0.000"

    ' Colonnes A a nbCols de la feuille brute, sans Date d'import ni Cellule source ;
    ' la valeur est reprise par formule * 1000.
    CreerCopie wsIdB, F_ID, Array("No enregistrement", "Societe", "Pays", "Branche", "Directeur general", _
        "Date de creation", "Capital social (F CFA)", "Cadres", "Maitrise", "Employes", "Annee N"), 11, ""
    CreerCopie wsEpB, F_EP, Array("No enregistrement", "Societe", "Pays", "Branche", "Bloc", "Categorie", _
        "Rubrique", "Mesure", "Annee", "Valeur (F CFA)"), 9, "J"
    CreerCopie wsCcB, NomChiffresCles(), Array("No enregistrement", "Societe", "Pays", "Branche", _
        "Rubrique", "Annee", "Valeur (F CFA)"), 6, "G"
End Sub

' Copie visible d'une feuille brute : colonnes A a nbCols en valeurs, puis,
' si colValeur est renseignee, une colonne Valeur (F CFA) = brute!colValeur * 1000.
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
        ws.Range("A2").Resize(derLig - 1, nbCols).Value = wsBrut.Range("A2").Resize(derLig - 1, nbCols).Value
        If Len(colValeur) > 0 Then
            With ws.Cells(2, nbCols + 1).Resize(derLig - 1, 1)
                .Formula = "=IF('" & wsBrut.Name & "'!" & colValeur & "2="""",""""," & _
                    "'" & wsBrut.Name & "'!" & colValeur & "2*1000)"
                .NumberFormat = "#,##0"
            End With
        End If
    End If

    If nom = F_ID Then
        ws.Columns("F").NumberFormat = "dd/mm/yyyy"
        ws.Columns("G").NumberFormat = "#,##0"
    End If
    ws.Visible = xlSheetVisible
    ws.Range("A1").Resize(1, nbEntetes).EntireColumn.AutoFit
End Sub

' ---------------------------------------------------------------------------------
' Conversions
' ---------------------------------------------------------------------------------

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
