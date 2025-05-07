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
            ' Call AddNewRow(wsTarget, colDict, Date, sourcePEP, sourceMaterial, "Jaraguá", 1320, sourceIncoterms)
        Else
            
            ' Check in wich column is the correct value for the following call
            If wsSource.Cells(i, sourceColDict("Wallet")).Value > wsSource.Cells(i, sourceColDict("Amount")).Value Then
                amount = wsSource.Cells(i, sourceColDict("Wallet")).Value
            Else
                amount = wsSource.Cells(i, sourceColDict("Amount")).Value
            End If
        
            ' Update the existing row at index j
            Call UpdateRowIfEmpty(wsTarget, j, colDict, Date, wsSource.Cells(i, sourceColDict("PEP")).Value, wsSource.Cells(i, sourceColDict("Market")).Value, wsSource.Cells(i, sourceColDict("Client")).Value, wsSource.Cells(i, sourceColDict("SalesDoc")).Value, "", "", wsSource.Cells(i, sourceColDict("Incoterms")).Value, wsSource.Cells(i, sourceColDict("Incoterms2")).Value, wsSource.Cells(i, sourceColDict("PM")).Value, amount, wsSource.Cells(i, sourceColDict("Plant")).Value)
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

Sub AddNewRow(wsTarget As Worksheet, colDict As Object, _
                   sourceDate As Date, sourcePEP As Variant, sourceMaterial As Variant, _
                   orderLocation As String, physicalStock As Variant, _
                   sourceIncoterms As Variant)
    Dim tbl As ListObject
    Dim newRow As ListRow
    ' Reference the table in wsTarget
    Set tbl = wsTarget.ListObjects("Tabela1")
    ' Add a new row to the table
    Set newRow = tbl.ListRows.Add
    
    ' Populate the new row with values
    newRow.Range.Cells(1, colDict("Date")).Value = sourceDate
    newRow.Range.Cells(1, colDict("PEP")).Value = sourcePEP
    newRow.Range.Cells(1, colDict("ZETO")).Value = sourceMaterial
    newRow.Range.Cells(1, colDict("ZVA1")).Value = sourceMaterial
    newRow.Range.Cells(1, colDict("OrderLocation")).Value = orderLocation
    newRow.Range.Cells(1, colDict("Incoterm")).Value = sourceIncoterms
    newRow.Range.Cells(1, colDict("StockStatus")).Value = "OK"
    newRow.Range.Cells(1, colDict("Checklist")).Value = "PENDENTE"
    newRow.Range.Cells(1, colDict("Freight")).Value = "PENDENTE"
    newRow.Range.Cells(1, colDict("Status")).Value = "AGUARD. PM"
    newRow.Range.Cells(1, colDict("PhysicalStock")).Value = physicalStock
    
End Sub

Sub UpdateRowIfEmpty(wsTarget As Worksheet, rowIndex As Long, colDict As Object, _
                     sourceDate As Date, sourcePEP As Variant, sourceMarket As Variant, sourceClient As Variant, _
                     sourceOV As Variant, sourceMaterial As Variant, sourceOrderLocation As String, _
                     sourceIncoterms As Variant, sourceIncoterms2 As Variant, sourcePM As Variant, _
                     sourceAmount As Variant, sourcePhysicalStock As Variant)
                
    With wsTarget
    
        ' Update Date if empty, Column A
        If False Then
            .Cells(rowIndex, colDict("Date")).Value = sourceDate
        End If
        
        ' Update PEP if empty, Column B
        If .Cells(rowIndex, colDict("PEP")).Value = "" Then
            .Cells(rowIndex, colDict("PEP")).Value = sourcePEP
        End If
        
        ' Update Market if empty, Column C
        If .Cells(rowIndex, colDict("Market")).Value = "" Then
            If InStr(1, UCase(sourceMarket), "FORA") > 0 Then
                .Cells(rowIndex, colDict("Market")).Value = "EXTERNO"
            Else
                .Cells(rowIndex, colDict("Market")).Value = "INTERNO"
            End If
        End If
        
        ' Update Client if empty, Column D
        If .Cells(rowIndex, colDict("Client")).Value = "" Then
            .Cells(rowIndex, colDict("Client")).Value = sourceClient
        End If
        
        ' Update OV if empty, Column E
        If .Cells(rowIndex, colDict("OV")).Value = "" Then
            .Cells(rowIndex, colDict("OV")).Value = sourceOV
        End If
        
        ' For ZETO (and ZVA1) update if both are empty, Column G
        If .Cells(rowIndex, colDict("ZETO")).Value <> "" And .Cells(rowIndex, colDict("ZVA1")).Value <> "" Then
            If .Cells(rowIndex, colDict("Market")).Value = "INTERNO" Then
                .Cells(rowIndex, colDict("ZETO")).Value = ""
                .Cells(rowIndex, colDict("ZVA1")).Value = sourceMaterial
            ElseIf .Cells(rowIndex, colDict("Market")).Value = "EXTERNO" Then
                .Cells(rowIndex, colDict("ZETO")).Value = sourceMaterial
                .Cells(rowIndex, colDict("ZVA1")).Value = ""
            End If
        End If
        
        ' Update OrderLocation if empty, Column I
        If .Cells(rowIndex, colDict("OrderLocation")).Value = "" Then
            .Cells(rowIndex, colDict("OrderLocation")).Value = sourceOrderLocation
        End If
        
        ' Update Incoterm if provided and the cell is empty, Column J
        If .Cells(rowIndex, colDict("Incoterm")).Value = "" Then
            .Cells(rowIndex, colDict("Incoterm")).Value = sourceIncoterms
        End If
        
        ' Update Incoterm2 if provided and the cell is empty, Column K
        If .Cells(rowIndex, colDict("Incoterm2")).Value = "" Then
            .Cells(rowIndex, colDict("Incoterm2")).Value = sourceIncoterms2
        End If
        
        ' Update PM if provided and the cell is empty, Column L
        If .Cells(rowIndex, colDict("PM")).Value = "" Then
            .Cells(rowIndex, colDict("PM")).Value = sourcePM
        End If
    
        ' Update Amount if provided and the cell is empty, Column M
        If .Cells(rowIndex, colDict("Amount")).Value = "" Then
            .Cells(rowIndex, colDict("Amount")).Value = sourceAmount
        End If
        
        ' Update PhysicalStock if empty, Column T
        If .Cells(rowIndex, colDict("PhysicalStock")).Value = "" Then
            .Cells(rowIndex, colDict("PhysicalStock")).Value = sourcePhysicalStock
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

