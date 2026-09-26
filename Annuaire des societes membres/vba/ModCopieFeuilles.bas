Attribute VB_Name = "ModCopieFeuilles"
Option Explicit

' Copie dans ce classeur toutes les feuilles des classeurs Excel du meme dossier.
'
' - Seul le contenu des cellules est copie : valeurs, formats et largeurs de colonnes.
'   Les formules deviennent des valeurs ; le code VBA des feuilles, les boutons, les noms
'   definis et les liaisons ne sont pas repris (un code VBA copie qui ne compile pas
'   bloquerait toutes les macros du classeur).
' - Les feuilles graphiques sont ignorees.
'
' - Les classeurs sont ouverts en lecture seule puis refermes sans modification.
' - Chaque feuille copiee est nommee "<nom du fichier> - <nom de la feuille>",
'   tronque a 31 caracteres (limite Excel) et rendu unique si besoin : "... (2)".
' - Ce classeur, les fichiers temporaires "~$..." et les feuilles tres masquees sont ignores.
' - Si le classeur est sur OneDrive / SharePoint (chemin en https://), Dir ne sait pas lire
'   le dossier : une fenetre demande alors de choisir le dossier local.
' - Un recapitulatif indique le nombre de feuilles copiees et les erreurs eventuelles.

Public Sub CopierFeuillesDuDossier()
    Dim dossier As String, fichier As String, erreurs As String
    Dim wbSrc As Workbook, sh As Object, fso As Object, f As Object
    Dim fichiers As New Collection, nom As Variant
    Dim nbFichiers As Long, nbFeuilles As Long
    Dim calcInitial As XlCalculation

    dossier = DossierDesFichiers()
    If Len(dossier) = 0 Then Exit Sub

    ' Liste des classeurs Excel du dossier (FileSystemObject gere les noms accentues)
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each f In fso.GetFolder(dossier).Files
        fichier = f.Name
        If LCase$(fso.GetExtensionName(fichier)) Like "xls*" _
           And StrComp(fichier, ThisWorkbook.Name, vbTextCompare) <> 0 _
           And Left$(fichier, 2) <> "~$" Then fichiers.Add fichier
    Next f
    If fichiers.Count = 0 Then
        MsgBox "Aucun autre classeur Excel dans le dossier :" & vbLf & dossier, vbInformation, "Copie"
        Exit Sub
    End If

    calcInitial = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    For Each nom In fichiers
        fichier = nom
        Set wbSrc = Nothing
        On Error Resume Next
        Set wbSrc = Workbooks.Open(Filename:=fso.BuildPath(dossier, fichier), UpdateLinks:=0, ReadOnly:=True)
        On Error GoTo 0

        If wbSrc Is Nothing Then
            erreurs = erreurs & vbLf & "- " & fichier & " : ouverture impossible"
        Else
            nbFichiers = nbFichiers + 1
            For Each sh In wbSrc.Worksheets
                If sh.Visible <> xlSheetVeryHidden Then
                    If CopierFeuille(sh, NomFichierSansExtension(fichier)) Then
                        nbFeuilles = nbFeuilles + 1
                    Else
                        erreurs = erreurs & vbLf & "- " & fichier & " / " & sh.Name & " : copie impossible"
                    End If
                End If
            Next sh
            Application.CutCopyMode = False
            wbSrc.Close SaveChanges:=False
        End If
    Next nom

    Application.Calculation = calcInitial
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox nbFeuilles & " feuille(s) copiee(s) depuis " & nbFichiers & " classeur(s)." & _
        IIf(Len(erreurs) > 0, vbLf & vbLf & "Problemes :" & erreurs, ""), vbInformation, "Copie"
End Sub

' Dossier de ce classeur s'il est lisible sur le disque ; sinon (classeur non enregistre,
' OneDrive / SharePoint en https://...) on demande le dossier. Renvoie "" si annule.
Private Function DossierDesFichiers() As String
    Dim fso As Object, p As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    p = ThisWorkbook.Path
    If Len(p) > 0 And LCase$(Left$(p, 4)) <> "http" Then
        If fso.FolderExists(p) Then
            DossierDesFichiers = p
            Exit Function
        End If
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Choisir le dossier des classeurs a copier"
        If .Show = -1 Then DossierDesFichiers = .SelectedItems(1)
    End With
End Function

' Cree une feuille a la fin de ce classeur et y colle le contenu de sh
' (formats, valeurs, largeurs de colonnes). Renvoie False en cas d'echec.
Private Function CopierFeuille(ByVal sh As Worksheet, ByVal prefixe As String) As Boolean
    Dim dest As Worksheet, zone As Range

    On Error GoTo Echec
    Set dest = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    dest.Name = NomUnique(prefixe & " - " & sh.Name)

    Set zone = sh.UsedRange
    zone.Copy
    With dest.Range(zone.Address)
        .PasteSpecial xlPasteColumnWidths
        .PasteSpecial xlPasteFormats
        .PasteSpecial xlPasteValues
    End With
    Application.CutCopyMode = False
    If sh.Visible = xlSheetHidden Then dest.Visible = xlSheetHidden

    CopierFeuille = True
    Exit Function

Echec:
    Application.CutCopyMode = False
    CopierFeuille = False
End Function

Private Function NomFichierSansExtension(ByVal fichier As String) As String
    Dim p As Long
    p = InStrRev(fichier, ".")
    If p > 0 Then NomFichierSansExtension = Left$(fichier, p - 1) Else NomFichierSansExtension = fichier
End Function

' Nom de feuille valide (sans : \ / ? * [ ], 31 caracteres max) et absent du classeur.
Private Function NomUnique(ByVal nom As String) As String
    Dim c As Variant, base As String, essai As String, i As Long

    For Each c In Array(":", "\", "/", "?", "*", "[", "]")
        nom = Replace(nom, c, "_")
    Next c
    nom = Trim$(nom)
    If Left$(nom, 1) = "'" Then nom = Mid$(nom, 2)
    If Right$(nom, 1) = "'" Then nom = Left$(nom, Len(nom) - 1)
    If Len(nom) = 0 Then nom = "Feuille"

    base = Left$(nom, 31)
    essai = base
    i = 1
    Do While FeuilleExiste(essai)
        i = i + 1
        essai = Left$(base, 31 - Len(" (" & i & ")")) & " (" & i & ")"
    Loop
    NomUnique = essai
End Function

Private Function FeuilleExiste(ByVal nom As String) As Boolean
    Dim sh As Object
    For Each sh In ThisWorkbook.Sheets
        If StrComp(sh.Name, nom, vbTextCompare) = 0 Then
            FeuilleExiste = True
            Exit Function
        End If
    Next sh
End Function
