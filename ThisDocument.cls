VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Dans Outils/Préférences activer les modules suivants :
' - Visual Basic for Application
' - Microsoft Word 15.0 Object Library
' - OLE Automation
' - Microsoft Office 15.0 Object Library
' - Microsoft Excel 15.0 Object Library
' - Microsoft Forms 2.0 Object Library
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Main()
    ' Modifier les noms des styles de titres pour qu'ils correspondent au document
    Const TITRE1_STYLE As String = "Titre 1"
    Const TITRE2_STYLE As String = "Titre 2"
    Const TITRE3_STYLE As String = "Titre 3"
    
    ' Modifier le nom du style de tableau d'exigence pour qu'ils correspondent au document
    Const EX_TABLE_STYLE As String = "Grille du tableau"
    
    'Nom du fichier excel qui contiendra l'extract des exigences. Il est créé au même endroit que les SFG.
    Const fileName As String = "SFG - Exigences.xlsx"
    
    'Debug.Print Selection.Range.Style
    Application.ScreenUpdating = False
    extractRequirements titre1Style:=TITRE1_STYLE, titre2Style:=TITRE2_STYLE, titre3Style:=TITRE3_STYLE, exTableStyle:=EX_TABLE_STYLE, fileName:=fileName
    Application.ScreenUpdating = True
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Extraction des exigences.
'
' La procédure parcours chaque paragraphe à la recherche des titres
' pour extraire le fonctionnel des SFG (qui est représenté par la
' hiérarchie des titres.
' Recherche également les tableaus dont la première cellule contient
' "ID" pour identifier les tableaux d'exigences.
' ATTENTION : aucun autre tableau ne doit contenir de cellule avec
' le seul terme "ID". C'est un point à améliorer.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub extractRequirements(titre1Style As String, titre2Style As String, titre3Style As String, exTableStyle As String, fileName As String)

' La nature d'un paragraphe est déterminée par son style :
' - Style de titre : fonctionnel ;
' - Style "Grille de tableau" : tableau d'exigence potentiel,
'                               confirmé si la première cellule
'                               contient exactement "ID".
Dim TITRE1_STYLE, TITRE2_STYLE, TITRE3_STYLE As String
TITRE1_STYLE = titre1Style
TITRE2_STYLE = titre2Style
TITRE3_STYLE = titre3Style

Dim title1Number, title1 As String
Dim title2Number, title2 As String
Dim title3Number, title3 As String
Dim hasNewTitle1, hasNewTitle2, hasNewTitle3 As Boolean
hasNewTitle1 = False
hasNewTitle2 = False
hasNewTitle3 = False

'Const TABLEAU_STYLE As String = "Grille du tableau"
Dim TABLEAU_STYLE As String
TABLEAU_STYLE = exTableStyle

'Style du paragraphe courant
Dim currentStyle As String

'Stockage d'une exigence
Dim exUID As String
Dim exText As String

' Création de la feuille excel
Dim objExcel As Excel.Application
Dim objDoc As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim toto As Word.Paragraph

Dim SaveAs1 As String
SaveAs1 = ActiveDocument.Path + "\" + fileName

Set objExcel = CreateObject("Excel.Application")
Set objDoc = objExcel.Workbooks.Add
objDoc.Sheets.Add Before:=objDoc.Sheets(1)
Set xlSheet = objDoc.Sheets(1)

' Mise en forme du fichier excel
setUpExigenceSheet xlSheet

objExcel.Visible = True

' On écrira les exigence à partir de la deuxième ligne
' car la première contient les entêtes
Dim xlRow As Integer
xlRow = 2


Debug.Print "#######################################"

For Each Para In ActiveDocument.Paragraphs
    If Not Para.Range.Style Is Nothing Then
    'Debug.Print para.Range.Style
    
    currentStyle = Para.Range.Style
    If currentStyle = TITRE1_STYLE Then
        'Debug.Print para.Range.ListFormat.ListString + " " + para.Range.Text
        hasNewTitle1 = True
        hasNewTitle2 = False
        hasNewTitle3 = False
        title1Number = Para.Range.ListFormat.ListString
        title1 = Para.Range.Text
        'insertAndStylesTitreInSheet xlRow:=xlRow, xlSheet:=xlSheet, titleNumber:=para.Range.ListFormat.ListString, title:=para.Range.Text, level:=0
        'xlRow = xlRow + 1
    ElseIf currentStyle = TITRE2_STYLE Then
        'Debug.Print para.Range.ListFormat.ListString + " " + para.Range.Text
        hasNewTitle2 = True
        hasNewTitle3 = False
        title2Number = Para.Range.ListFormat.ListString
        title2 = Para.Range.Text
        'insertAndStylesTitreInSheet xlRow:=xlRow, xlSheet:=xlSheet, titleNumber:=para.Range.ListFormat.ListString, title:=para.Range.Text, level:=1
        'xlRow = xlRow + 1
    ElseIf currentStyle = TITRE3_STYLE Then
        'Debug.Print para.Range.ListFormat.ListString + " " + para.Range.Text
        hasNewTitle3 = True
        title3Number = Para.Range.ListFormat.ListString
        title3 = Para.Range.Text
        'insertAndStylesTitreInSheet xlRow:=xlRow, xlSheet:=xlSheet, titleNumber:=para.Range.ListFormat.ListString, title:=para.Range.Text, level:=2
        'xlRow = xlRow + 1
    Else
            If currentStyle = TABLEAU_STYLE Then
            Para.Range.Select
            If isExigenceTable() Then
                If hasNewTitle1 Then
                    insertAndStylesTitreInSheet xlRow:=xlRow, xlSheet:=xlSheet, titleNumber:=CStr(title1Number), title:=title1, level:=0
                    hasNewTitle1 = False
                    xlRow = xlRow + 1
                End If
                If hasNewTitle2 Then
                    insertAndStylesTitreInSheet xlRow:=xlRow, xlSheet:=xlSheet, titleNumber:=CStr(title2Number), title:=title2, level:=1
                    hasNewTitle2 = False
                    xlRow = xlRow + 1
                End If
                If hasNewTitle3 Then
                    insertAndStylesTitreInSheet xlRow:=xlRow, xlSheet:=xlSheet, titleNumber:=CStr(title3Number), title:=title3, level:=2
                    hasNewTitle3 = False
                    xlRow = xlRow + 1
                End If
                For iRow = 2 To Selection.Tables(1).Rows.Count
                    exUID = Selection.Tables(1).Cell(iRow, 1).Range.Text
                    exUID = Mid(exUID, 1, Len(exUID) - 2)
                    exText = Selection.Tables(1).Cell(iRow, 2).Range.Text
                    exText = Mid(exText, 1, Len(exText) - 2)
                    'Debug.Print exUID + " " + exText
                    xlSheet.Cells(xlRow, 1) = exUID
                    xlSheet.Cells(xlRow, 2) = exText
                    xlRow = xlRow + 1
                Next
            End If
        End If
    End If
    End If
Next

objExcel.ActiveWorkbook.SaveAs fileName:=SaveAs1, CreateBackup:=False


objDoc.Close
objExcel.Quit


End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Détermine si la sélection courante est la première cellule d'un
' tableau d'exigences.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function isExigenceTable() As Boolean
    Dim ret As Boolean
        
    With Selection.Find
                .Forward = True
                .ClearFormatting
                .MatchWholeWord = True
                .MatchCase = True
                .Wrap = wdFindStop 'wdFindContinue
                .Text = "ID"
                isExigenceTable = .Execute
    End With
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialisation et mise en forme de la feuille excel qui va
' contenir la liste des exigences
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub setUpExigenceSheet(xlSheet As Excel.Worksheet)
xlSheet.Name = "Exigences"
With xlSheet.Rows(1)
    .Font.Bold = True
    .Interior.Color = RGB(217, 217, 217)
    .VerticalAlignment = xlCenter
End With


With xlSheet.Cells(1, 1)
 .Value = "ID"
 .ColumnWidth = 16
End With

With xlSheet.Cells(1, 2)
    .Value = "Exigence"
    .ColumnWidth = 82
End With

With xlSheet.Cells(1, 3)
    .Value = "Paramétrage" & Chr(10) & "JH"
    .ColumnWidth = 12
    .HorizontalAlignment = xlCenter
End With

With xlSheet.Cells(1, 4)
    .Value = "Dev Spécifique" & Chr(10) & "JH"
    .ColumnWidth = 15
    .HorizontalAlignment = xlCenter
End With

With xlSheet.Cells(1, 5)
    .Value = "Total" & Chr(10) & "JH"
    .ColumnWidth = 12
    .HorizontalAlignment = xlCenter
End With

With xlSheet.Cells(1, 6)
    .Value = "Hypothèses de chiffrage"
    .ColumnWidth = 57
End With

xlSheet.Range("A:A").Font.Bold = True
xlSheet.Range("B:B").WrapText = True
xlSheet.Range("A2:F500").VerticalAlignment = xlTop
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Création de la ligne d'entête du fichier excel qui va contenir
' les exigences.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub insertAndStylesTitreInSheet(xlRow As Integer, xlSheet As Excel.Worksheet, titleNumber As String, title As String, level As Integer)
    Dim rowColor
    xlSheet.Cells(xlRow, 1) = titleNumber + " " + title
    xlSheet.Cells(xlRow, 1).IndentLevel = level
    
    Select Case level
        Case 0
        rowColor = RGB(155, 194, 230)
        Case 1
        rowColor = RGB(189, 215, 238)
        Case Else
        rowColor = RGB(221, 235, 247)
    End Select
        
    xlSheet.Rows(xlRow).Interior.Color = rowColor
End Sub


