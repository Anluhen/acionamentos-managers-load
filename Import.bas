Attribute VB_Name = "Import"
Option Explicit

Public SapGuiAuto As Object
Public SAPApplication As Object
Public Connection As Object
Public session As Object

Sub AtualizarDados()
    
    On Error GoTo ErrorHandler
    
    OptimizeCodeExecution True
    
    Dim wbThis As Workbook, exportWb As Workbook, wb As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet
    'Dim sourceLastRow As Long, sourceLastCol As Long, targetLastRow As Long, targetLastCol As Long
    'Dim sourceHeaderRow As Range
    'Dim sourceColIndex As Long, targetColIndex As Long
    'Dim sourcePEP As String, sourceMaterial As String, sourceValor As String, sourceIncoterms As String
    'Dim targetPEP As String, targetZETO As String, targetZVA1 As String
    'Dim i As Long, j As Long
    'Dim cell As Range
    'Dim exportWbName As String, exportWbPath As String
    'Dim tries As Long
    'Dim found As Boolean, isNotFound As Boolean
    'Dim ErrSection As String
    'Dim sourceColDict As Object
    'Dim startTime As Double
    'Dim key As Variant
    'Dim amount As Double
    
ErrSection = "variableDeclarations"
    
    Set wbThis = ThisWorkbook
    
    Set wsTarget = wbThis.Sheets("PEDIDOS 2025")
    
ErrSection = "variableDeclarations10"

    If wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If
    
    ' Clear all filters if any
    If wsTarget.ListObjects(1).ShowAutoFilter Then
        wsTarget.ListObjects(1).AutoFilter.ShowAllData
    End If
    
    Dim targetColMap As Object
    Set targetColMap = MapTargetColumnHeaders()
    
    ' Find the last used row and column in the source sheet
    targetLastRow = wsTarget.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    targetLastCol = wsTarget.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "OpenSource"

    ' Open the workbook Reunião de Faturamento Semanal - New Layout.xlsm
    Workbooks.Open "https://weg365.sharepoint.com/teams/BR-WAU-VENDAS-ADMCONTRATOS/Shared%20Documents/REUNI%C3%83O%20DE%20FATURAMENTO/Reuni%C3%A3o%20de%20Faturamento%20Semanal%20-%20New%20Layout.xlsm"
    exportWbName = "Reunião de Faturamento Semanal - New Layout"
    
    tries = 0

    Do
        If tries > 10 Then
            GoTo ErrorHandler
        End If
        
        found = False
        
        ' Loop through all open workbooks
        For Each wb In Application.Workbooks
            If UCase(wb.Name) = UCase(exportWbName) Then
                Set exportWb = wb
                found = True
                Exit Do
            End If
        Next wb
    
        tries = tries + 1
        
        DoEvents
    Loop

ErrSection = "SourceSheet"

    ' Set worksheets
    Set wsSource = exportWb.Sheets("Faturamento")

     If wsSource Is Nothing Or wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If

ErrSection = "completeInformationFromAnalisys20"
    
    ' Find the last used row and column in the source sheet
    sourceLastRow = wsSource.Cells(wsSource.Rows.Count, 2).End(xlUp).Row
    sourceLastCol = wsSource.Cells(2, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "completeInformationFromAnalisys30"

    Dim sourceColMap As Object
    Set sourceColMap = MapSourceColumnHeaders()
    
    ' Loop through rows from bottom to top to avoid skipping rows after deletion
    For i = sourceLastRow To 2 Step -1 ' Assuming headers are in row 1
ErrSection = "completeInformationFromAnalisys40-" & i
        isNotFound = True

        If InStr(wsSource.Cells(i, sourceColDict("PEP")).Value, "-") <> 0 Then
            For j = targetLastRow To 3 Step -1 ' Assuming headers are in row 1
                If Left(wsSource.Cells(i, sourceColDict("PEP")).Value, InStr(InStr(wsSource.Cells(i, sourceColDict("PEP")).Value, "-") + 1, wsSource.Cells(i, sourceColDict("PEP")).Value, "-") - 1) = Left(wsTarget.Cells(j, colDict("PEP")).Value, InStr(InStr(wsTarget.Cells(j, colDict("PEP")).Value, "-") + 1, wsTarget.Cells(j, colDict("PEP")).Value, "-") - 1) Then
                    isNotFound = False
                    Exit For
                End If
            Next j
        End If
        
        If isNotFound Then
            ' Create a new row with OrderLocation "Jaraguá", PhysicalStock 1320 and include Incoterm
            Call AddNewRowFromAnalysis(wsTarget, targetColMap, wsSource.Cells(i, sourceColDict("PEP")).Row, sourceColMap)
        Else
            ' Update the existing row at index j
            Call UpdateRowFromAnalysis(wsTarget, targetColMap, j, wsSource.Cells(i, sourceColDict("PEP")).Row, sourceColMap)
        End If
    Next i

completeLocationInfo:
ErrSection = "completeLocationInfo"
    
    Set wsTarget = wbThis.Sheets("FATURAMENTO")
    
     If wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If

ErrSection = "completeLocationInfo10"

    ' Find the last used row and column in the source sheet
    targetLastRow = wsTarget.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    targetLastCol = wsTarget.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "completeLocationInfo20"

    ' If "Situação" column not found, exit sub
    If colDict("Status") = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Loop through rows from bottom to top to avoid skipping rows after deletion
    For i = targetLastRow To 2 Step -1 ' Assuming headers are in row 1
ErrSection = "completeLocationInfo30-" & i
        With wsTarget.Cells(i, colDict("PhysicalStock"))
            If wsTarget.Cells(i, colDict("OrderLocation")).Value = "" Then
                ' Trim to first 4 characters if longer than 4
                If Len(.Value) > 4 Then
                    .Value = Left(.Value, 4)
                End If
                
                ' Assign OrderLocation based on PhysicalStock
                Select Case .Value
                    Case "1320"
                        wsTarget.Cells(i, colDict("OrderLocation")).Value = "JGS"
                    Case "1321"
                        wsTarget.Cells(i, colDict("OrderLocation")).Value = "ITJ"
                End Select
            End If
            
            If wsTarget.Cells(i, colDict("PhysicalStock")).Value = "" Then
                ' Assign PhysicalStock based on OrderLocation
                If InStr(1, UCase(wsTarget.Cells(i, colDict("OrderLocation")).Value), "JGS") > 0 Then
                    wsTarget.Cells(i, colDict("PhysicalStock")).Value = 1320
                ElseIf InStr(1, UCase(wsTarget.Cells(i, colDict("OrderLocation")).Value), "ITJ") > 0 Then
                    wsTarget.Cells(i, colDict("PhysicalStock")).Value = 1321
                End If
            End If
        End With
    Next i
    
    ' Success message
    MsgBox "Os dados foram atualizados com sucesso!", vbInformation, "Macro Finalizada"

' Ignore next bit of code
GoTo CleanExit
ErrSection = "moveFinishedItems"

    ' Set worksheets
    Set wsSource = wbThis.Sheets("FATURAMENTO")
    Set wsTarget = wbThis.Sheets("Finalizado")
    
ErrSection = "moveFinishedItems10"

    If wsSource Is Nothing Or wsTarget Is Nothing Then
        GoTo ErrorHandler
    End If

    ' Find the last used row and column in the source sheet
    sourceLastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    sourceLastCol = wsSource.Cells(1, wsSource.Columns.Count).End(xlToLeft).Column

ErrSection = "moveFinishedItems20"

    ' If "Situação" column not found, exit sub
    If colDict("Status") = 0 Then
        GoTo ErrorHandler
    End If
    
    ' Loop through rows from bottom to top to avoid skipping rows after deletion
    For i = sourceLastRow To 2 Step -1 ' Assuming headers are in row 1
ErrSection = "moveFinishedItems30-" & i
        If UCase(Trim(wsSource.Cells(i, colDict("Status")).Value)) = "RECONHECIDO" Then
            ' Find last row in target sheet
            targetLastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1

            ' Copy and paste formats from the row above
            wsTarget.Rows(targetLastRow - 1).Copy
            wsTarget.Rows(targetLastRow).PasteSpecial Paste:=xlFormats
            
            ' Copy and paste values from the source
            wsSource.Rows(i).Copy
            wsTarget.Rows(targetLastRow).PasteSpecial Paste:=xlValue
            
            ' Clear clipboard
            Application.CutCopyMode = False
            
            ' Delete the original row to avoid empty rows
            wsSource.Rows(i).Delete Shift:=xlUp
        End If
    Next i

    ' Clear clipboard
    Application.CutCopyMode = False

CleanExit:

    ' Loop through all open workbooks to find if the exportWb is oppened
    For Each wb In Application.Workbooks
        If UCase(wb.Name) = UCase(exportWbName) Then
            wb.Close False ' Close the exportWb
        End If
    Next wb
    
    ' Ensure that all optimizations are turned back on
    OptimizeCodeExecution False
    Exit Sub

ErrorHandler:
        
    ' Log and diagnose the error using Erl to show the last executed line number
    MsgBox "Erro " & Err.Number & " após " & ErrSection & ": " & Err.Description, vbCritical, "Erro em AtualizarDados"
    
    ' Resume cleanup to ensure that settings are restored
    GoTo CleanExit
End Sub

Sub AddNewRowFromAnalysis(wsTarget As Worksheet, targetColMap As Object, wsSourceRow As Range, sourceColMap As Object)
    Dim tbl As ListObject
    Dim newRow As ListRow
    ' Reference the table in wsTarget
    Set tbl = wsTarget.ListObjects("Tabela1")
    ' Add a new row to the table
    Set newRow = tbl.ListRows.Add
    
    ' Populate the new row with values
    ' newRow.Range.Cells(1, targetColMap("DATA")).Value = ""
    ' newRow.Range.Cells(1, targetColMap("ANO")).Value = ""
    ' newRow.Range.Cells(1, targetColMap("MÊS")).Value = ""
    ' newRow.Range.Cells(1, targetColMap("NOTA")).Value = ""
    ' newRow.Range.Cells(1, targetColMap("DATA FIM")).Value = ""
    ' newRow.Range.Cells(1, targetColMap("DATA DOC.")).Value = ""
    newRow.Range.Cells(1, targetColMap("ORDEM DE VENDA")).Value = wsSourceRow.Cells(1, sourceColMap("Doc. Vendas")).Value
    newRow.Range.Cells(1, targetColMap("DATA PREP")).Value = wsSourceRow.Cells(1, sourceColMap("Data Prep.")).Value
    newRow.Range.Cells(1, targetColMap("VALOR (BRL)")).Value = wsSourceRow.Cells(1, sourceColMap("Valor")).Value
    newRow.Range.Cells(1, targetColMap("CLIENTE")).Value = wsSourceRow.Cells(1, sourceColMap("Cliente")).Value
    newRow.Range.Cells(1, targetColMap("PEP")).Value = wsSourceRow.Cells(1, sourceColMap("PEP")).Value
    ' newRow.Range.Cells(1, targetColMap("SCORECARD")).Value = ""
    ' newRow.Range.Cells(1, targetColMap("Antecipação")).Value = ""
    newRow.Range.Cells(1, targetColMap("PM")).Value = wsSourceRow.Cells(1, sourceColMap("PM")).Value
    newRow.Range.Cells(1, targetColMap("Status")).Value = "ADD. MACRO"

End Sub

Sub UpdateRowFromAnalysis(wsTarget As Worksheet, targetColMap As Object, targetRowIndex As Long, wsSourceRow As Range, sourceColMap As Object)
                
    With wsTarget
    
        If .Cells(targetRowIndex, targetColMap("DATA")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("DATA")).Value = ""
        End If

        If .Cells(targetRowIndex, targetColMap("ANO")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("ANO")).Value = ""
        End If

        If .Cells(targetRowIndex, targetColMap("MÊS")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("MÊS")).Value = ""
        End If

        If .Cells(targetRowIndex, targetColMap("NOTA")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("NOTA")).Value = ""
        End If

        If .Cells(targetRowIndex, targetColMap("DATA FIM")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("DATA FIM")).Value = ""
        End If

        If .Cells(targetRowIndex, targetColMap("DATA DOC.")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("DATA DOC.")).Value = ""
        End If

        If .Cells(targetRowIndex, targetColMap("ORDEM DE VENDA")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("ORDEM DE VENDA")).Value = wsSourceRow.Cells(targetRowIndex, sourceColMap("Doc. Vendas")).Value
        End If

        If .Cells(targetRowIndex, targetColMap("DATA PREP")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("DATA PREP")).Value = wsSourceRow.Cells(targetRowIndex, sourceColMap("Data Prep.")).Value
        End If

        If .Cells(targetRowIndex, targetColMap("VALOR (BRL)")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("VALOR (BRL)")).Value = wsSourceRow.Cells(targetRowIndex, sourceColMap("Valor")).Value
        End If

        If .Cells(targetRowIndex, targetColMap("CLIENTE")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("CLIENTE")).Value = wsSourceRow.Cells(targetRowIndex, sourceColMap("Cliente")).Value
        End If

        If .Cells(targetRowIndex, targetColMap("PEP")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("PEP")).Value = wsSourceRow.Cells(targetRowIndex, sourceColMap("PEP")).Value
        End If

        If .Cells(targetRowIndex, targetColMap("SCORECARD")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("SCORECARD")).Value = ""
        End If

        If .Cells(targetRowIndex, targetColMap("Antecipação")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("Antecipação")).Value = ""
        End If

        If .Cells(targetRowIndex, targetColMap("PM")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("PM")).Value = wsSourceRow.Cells(targetRowIndex, sourceColMap("PM")).Value
        End If

        If .Cells(targetRowIndex, targetColMap("Status")).Value = "" Then
            .Cells(targetRowIndex, targetColMap("Status")).Value = "UPD. MACRO"
        End If
    End With
End Sub

Function GetSourceColumnIndex(ws As Worksheet, headerName As String, headerRow As Long, sourceColDict As Object) As Long
    Dim col As Range
    Dim alreadyUsed As Boolean
    
    GetSourceColumnIndex = 0 ' Not found
    
    For Each col In ws.Rows(headerRow).Cells
        If InStr(1, Trim(UCase(col.Value)), Trim(UCase(headerName))) > 0 Then
            alreadyUsed = False

            ' Check if this column number is already assigned in colDict
            Dim key As Variant
            For Each key In sourceColDict.Keys
                If sourceColDict(key) = col.Column Then
                    alreadyUsed = True
                    Exit For
                End If
            Next key

            ' If not already used, assign it
            If Not alreadyUsed Then
                GetSourceColumnIndex = col.Column
                Exit Function
            End If
        End If
    Next col
End Function

Public Function MapTargetColumnHeaders() As Object
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    headers.Add "DATA", 1
    headers.Add "ANO", 2
    headers.Add "MÊS", 3
    headers.Add "NOTA", 4
    headers.Add "DATA FIM", 5
    headers.Add "DATA DOC.", 6
    headers.Add "ORDEM DE VENDA", 7
    headers.Add "DATA PREP", 8
    headers.Add "VALOR (BRL)", 9
    headers.Add "CLIENTE", 10
    headers.Add "PEP", 11
    headers.Add "SCORECARD", 12
    headers.Add "Antecipação", 13
    headers.Add "PM", 14
    headers.Add "Status", 15

    Set MapTargetColumnHeaders = headers
End Function

Public Function MapSourceColumnHeaders() As Object
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    
    headers.Add "ID", 1                          ' Column B
    headers.Add "Mercado", 2                     ' Column C
    headers.Add "Ano BI", 3                      ' Column D
    headers.Add "Mês BI", 4                      ' Column E
    headers.Add "STATUS", 5                      ' Column F
    headers.Add "Doc. Compra", 6                 ' Column G
    headers.Add "Doc. Vendas", 7                 ' Column H
    headers.Add "Item Doc. Venda", 8             ' Column I
    headers.Add "PEP", 9                         ' Column J
    headers.Add "Cliente", 10                    ' Column K
    headers.Add "Incoterms", 11                  ' Column L
    headers.Add "Data Prep. Material", 12        ' Column M
    headers.Add "Data Remessa", 13               ' Column N
    headers.Add "Data Adc. B ", 14               ' Column O
    headers.Add "Data NF ", 15                   ' Column P
    headers.Add "Data PCP", 16                   ' Column Q
    headers.Add "Data de Rec.Receita ", 17       ' Column R
    headers.Add "Área Causadora", 18             ' Column S
    headers.Add "Motivo", 19                     ' Column T
    headers.Add "Observação", 20                 ' Column U
    headers.Add "Situação", 21                   ' Column V
    headers.Add "PM", 22                         ' Column W
    headers.Add "Valor", 23                      ' Column X
    headers.Add "Centro", 24                     ' Column Y
    headers.Add "Hier. Produto", 25              ' Column Z
    headers.Add "Mês", 26                        ' Column AA
    headers.Add "Ano", 27                        ' Column AB

    Set MapSourceColumnHeaders = headers
End Function

Function GetColumnIndex(ws As Worksheet, headerName As String, Optional headerRow As Long = 1) As Long
    Dim col As Range
    Dim alreadyUsed As Boolean
    
    GetColumnIndex = 0 ' Not found
    
    For Each col In ws.Rows(headerRow).Cells
        If InStr(1, Trim(UCase(col.Value)), Trim(UCase(headerName))) > 0 Then
            alreadyUsed = False

            ' Check if this column number is already assigned in colDict
            Dim key As Variant
            For Each key In colDict.Keys
                If colDict(key) = col.Column Then
                    alreadyUsed = True
                    Exit For
                End If
            Next key

            ' If not already used, assign it
            If Not alreadyUsed Then
                GetColumnIndex = col.Column
                Exit Function
            End If
        End If
    Next col
End Function

Function OptimizeCodeExecution(enable As Boolean)
    With Application
        If enable Then
            ' Disable settings for optimization
            .ScreenUpdating = False
            .Calculation = xlCalculationManual
            .EnableEvents = False
        Else
            ' Re-enable settings after optimization
            .ScreenUpdating = True
            .Calculation = xlCalculationAutomatic
            .EnableEvents = True
        End If
    End With
End Function

