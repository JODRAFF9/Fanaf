Attribute VB_Name = "ModCopieFeuilles"
Option Explicit

' Copie dans ce classeur toutes les feuilles des classeurs Excel du meme dossier.
'
' - Les classeurs sont ouverts en lecture seule puis refermes sans modification.
' - Chaque feuille copiee est nommee "<nom du fichier> - <nom de la feuille>",
'   tronque a 31 caracteres (limite Excel) et rendu unique si besoin : "... (2)".
' - Ce classeur, les fichiers temporaires "~$..." et les feuilles tres masquees sont ignores.
' - Un recapitulatif indique le nombre de feuilles copiees et les erreurs eventuelles.

Public Sub CopierFeuillesDuDossier()
    Dim dossier As String, fichier As String, erreurs As String
    Dim wbSrc As Workbook, sh As Object
    Dim nbFichiers As Long, nbFeuilles As Long
    Dim calcInitial As XlCalculation

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Enregistrez d'abord ce classeur dans le dossier des fichiers a copier.", vbExclamation, "Copie"
        Exit Sub
    End If
    dossier = ThisWorkbook.Path & Application.PathSeparator

    calcInitial = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    fichier = Dir(dossier & "*.xls*")
    Do While Len(fichier) > 0
        If fichier <> ThisWorkbook.Name And Left$(fichier, 2) <> "~$" Then
            Set wbSrc = Nothing
            On Error Resume Next
            Set wbSrc = Workbooks.Open(Filename:=dossier & fichier, UpdateLinks:=0, ReadOnly:=True)
            On Error GoTo 0

            If wbSrc Is Nothing Then
                erreurs = erreurs & vbLf & "- " & fichier & " : ouverture impossible"
            Else
                nbFichiers = nbFichiers + 1
                For Each sh In wbSrc.Sheets
                    If sh.Visible <> xlSheetVeryHidden Then
                        If CopierFeuille(sh, NomFichierSansExtension(fichier)) Then
                            nbFeuilles = nbFeuilles + 1
                        Else
                            erreurs = erreurs & vbLf & "- " & fichier & " / " & sh.Name & " : copie impossible"
                        End If
                    End If
                Next sh
                wbSrc.Close SaveChanges:=False
            End If
        End If
        fichier = Dir()
    Loop

    Application.Calculation = calcInitial
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox nbFeuilles & " feuille(s) copiee(s) depuis " & nbFichiers & " classeur(s)." & _
        IIf(Len(erreurs) > 0, vbLf & vbLf & "Problemes :" & erreurs, ""), vbInformation, "Copie"
End Sub

' Copie une feuille a la fin de ce classeur et la renomme. Renvoie False en cas d'echec.
Private Function CopierFeuille(ByVal sh As Object, ByVal prefixe As String) As Boolean
    Dim nb As Long

    On Error GoTo Echec
    nb = ThisWorkbook.Sheets.Count
    sh.Copy After:=ThisWorkbook.Sheets(nb)
    ThisWorkbook.Sheets(nb + 1).Name = NomUnique(prefixe & " - " & sh.Name)
    CopierFeuille = True
    Exit Function

Echec:
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
