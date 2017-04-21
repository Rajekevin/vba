
Sub ExtractDate()
Dim rng As Range, Dn As Range, Sp As Variant, n As Long
Set rng = Range(Range("A2"), Range("A" & Rows.Count).End(xlUp))
For Each Dn In rng
    Sp = Split(Dn.Value, " ")
    For n = 0 To UBound(Sp)
        If IsDate(Sp(n)) Then Dn.Offset(, 2) = Sp(n)
    Next n
Next Dn
End Sub

Sub ExtractCivilite()
Dim rng As Range, Dn As Range
Set rng = Range(Range("A2"), Range("A" & Rows.Count).End(xlUp))
For Each Dn In rng
    'replace(Dn.Value, "M. ", "")
    Dn = Replace(Dn, "M. ", "")
    Dn = Replace(Dn, "Mme ", "")
    Dn = Trim(Dn)
Next Dn
End Sub

Sub ExtractNomPrenom()
Dim rng As Range, Dn As Range, Sp As Variant, NP As Variant, n As Long
Set rng = Range(Range("A2"), Range("A" & Rows.Count).End(xlUp))
For Each Dn In rng
    Sp = Split(Dn.Value, " né(e) le ")
    NP = Split(Sp(0), " ")
    'MsgBox (Dn.Row)
    Range("A" & Dn.Row) = NP(0)
    Range("B" & Dn.Row) = NP(1)
    
Next Dn
End Sub




Sub Export_CSV()

    Dim MyPath As String
    Dim MyFileName As String
    Dim WB1 As Workbook, WB2 As Workbook
     
    Set WB1 = ActiveWorkbook

    Dim rng As Range
    Set rng = Application.InputBox("select cell range with changes", "Cells to be copied", Default:="Select Cell Range", Type:=8)
    Application.ScreenUpdating = False
    rng.Copy
 
    Set WB2 = Application.Workbooks.Add(1)
    WB2.Sheets(1).Range("A1").PasteSpecial xlPasteValues
     
    MyFileName = "CSV_Export_" & Format(Date, "ddmmyyyy")
    FullPath = WB1.Path & "\" & MyFileName
     
    Application.DisplayAlerts = False
    If MsgBox("Data copied to " & WB1.Path & "\" & MyFileName & vbCrLf & _
    "Attention: Les fichiers dans le dossier avec le même nom sera écrasé", vbQuestion + vbYesNo) <> vbYes Then
        Exit Sub
    End If
     
    If Not Right(MyFileName, 4) = ".csv" Then MyFileName = MyFileName & ".csv"
    With WB2
        .SaveAs Filename:=FullPath, FileFormat:=xlCSV, CreateBackup:=False
        .Close False
    End With
    Application.DisplayAlerts = True
End Sub

Sub Header()


'Range("A1:M1").Interior.Color = RGB(66, 185, 244)  'Background Color
'Range("A1:M1").Font.Bold = True                    'Bold
'Range("A1:M1").Font.Italic = True                  'Italic
'Range("A1:M1").Font.Color = RGB(255, 255, 255)     'Header color white

    For Each ws In Worksheets

    'ws.Range("F1").EntireColumn.Hidden = True
    'ws.Range("D1").EntireColumn.Hidden = True
    ws.Rows(1).Insert
    ws.Range("A1:E1") = Array("Nom", "Prenom", "Date de Naissance", "Ouverture", "Statut")
    ws.Range("A1:E1").Interior.Color = RGB(66, 185, 244)
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A1:E1").Font.Italic = True
    ws.Range("A1:E1").Font.Color = RGB(255, 255, 255)
    
    
Next ws
End Sub


Sub ExtractNomPrenom1()
  Call ExtractDate
  Call ExtractCivilite
  Call ExtractNomPrenom

End Sub

'Nettoie le fichier et l'export en CSV
Sub MajCsv()
  Dim rng As Range
  Call Header
  Call ExtractDate
  Call ExtractCivilite
  Call ExtractNomPrenom
  Call Export_CSV
End Sub



Sub Title()
    Rows(2).EntireRow.Delete
    Range("A2").EntireRow.Insert
    Range("A1:M1").Interior.Color = RGB(66, 185, 244)  'Background Color
    Range("A1:M1").Font.Bold = True                    'Bold
    Range("A1:M1").Font.Italic = True                  'Italic
    Range("A1:M1").Font.Color = RGB(255, 255, 255)     'Header color white
    
    Range("A1").Value = "Nom"
    Range("B1").Value = "Prenom"
    Range("C1").Value = "Date de Naissance"
    Range("D1").Value = "Ouverture"
    Range("E1").Value = "Statut"

End Sub
