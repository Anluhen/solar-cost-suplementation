Attribute VB_Name = "Custos"
' ----- Version -----
'        1.2.0
' -------------------

Sub SaveData(Optional ShowOnMacroList As Boolean = False)

    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim tblRow As ListRow
    Dim newID As String
    Dim userResponse As VbMsgBoxResult
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    newID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' If ComboBoxID is not empty, prompt the user
    If Trim(newID) <> "" Then
        userResponse = MsgBox("Esse aditivo já foi cadastrado. Deseja sobrescrever?", vbYesNoCancel + vbQuestion, "Confirmação")

        Select Case userResponse
            Case vbYes
                newID = Val(newID) ' Use ComboBoxID.Value as new ID
                ' Search for the ID in the first column of the table
                Set tblRow = dadosTable.ListRows(dadosTable.ListColumns(1).DataBodyRange.Find(What:=newID, LookAt:=xlWhole).Row - dadosTable.DataBodyRange.Row + 1)
            Case vbNo
                ' Proceed with generating new ID
                newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(1).DataBodyRange) + 1
                wsForm.OLEObjects("ComboBoxID").Object.Value = newID
                ' Add a new row to the table
                Set tblRow = dadosTable.ListRows.Add
            Case vbCancel
                Exit Sub ' Exit without saving
        End Select
    Else
        newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(1).DataBodyRange) + 1
        
        wsForm.OLEObjects("ComboBoxID").Object.Value = newID
        
        wsForm.OLEObjects("ComboBoxName").Object.Value = wsForm.Range("B6").Value & " - " & wsForm.Range("B10").Value & " - " & wsForm.Range("D6").Value
        
        ' Add a new row to the table
        Set tblRow = dadosTable.ListRows.Add
    End If
    
    ' Assign values to the new row
    With tblRow.Range
        ' Set new ID
        .Cells(1, 1).Value = newID ' First column value
        
        ' Read column B values
        .Cells(1, 2).Value = wsForm.Range("B6").Value
        .Cells(1, 3).Value = wsForm.Range("B10").Value
        .Cells(1, 4).Value = wsForm.Range("B14").Value
        .Cells(1, 5).Value = wsForm.Range("B18").Value
        .Cells(1, 6).Value = wsForm.Range("B22").Value
        .Cells(1, 7).Value = wsForm.Range("B28").Value
        .Cells(1, 8).Value = wsForm.Range("B32").Value
        .Cells(1, 9).Value = wsForm.Range("B36").Value
        .Cells(1, 10).Value = wsForm.Range("B40").Value
        .Cells(1, 11).Value = wsForm.Range("B44").Value
        .Cells(1, 12).Value = wsForm.Range("B48").Value
        .Cells(1, 13).Value = wsForm.Range("B52").Value
        
        If .Cells(1, 14).Formula = "" Then
            .Cells(1, 14).Formula = "=IFERROR([@[Custo da Suplementação]]/[@[Custo COT]];"")"
        End If
        If .Cells(1, 15).Formula = "" Then
            .Cells(1, 15).Formula = "=[@[Custo Planejado Depois da Suplementação]]-[@[Custo COT]]-[@[Custo da Suplementação]]"
        End If
        
        .Cells(1, 16).Value = wsForm.Range("B56").Value
        
        ' Read column D values
        .Cells(1, 17).Value = wsForm.Range("D6").Value
        .Cells(1, 18).Value = wsForm.Range("D10").Value
        
        If wsForm.Range("D14").Value < 0.4 Then
            .Cells(1, 19).Value = Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Inicial)"
        ElseIf wsForm.Range("D14").Value < 0.8 Then
            .Cells(1, 19).Value = Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Intermediária)"
        Else
            .Cells(1, 19).Value = Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Final)"
        End If

        .Cells(1, 20).Value = wsForm.Range("D18").Value
        .Cells(1, 21).Value = wsForm.Range("D22").Value
        .Cells(1, 22).Value = wsForm.Range("D26").Value
        .Cells(1, 23).Value = wsForm.Range("D30").Value
        .Cells(1, 24).Value = wsForm.Range("D34").Value
        .Cells(1, 25).Value = wsForm.Range("D38").Value
    
        ' Read column F values
        .Cells(1, 26).Value = wsForm.Range("F6").Value
        .Cells(1, 27).Value = wsForm.Range("F10").Value
        .Cells(1, 28).Value = wsForm.Range("F14").Value
        .Cells(1, 29).Value = wsForm.Range("F18").Value
        .Cells(1, 30).Value = "" ' Erase date if information is overwriten to allow resend the e-mail
        .Cells(1, 31).Value = wsForm.Range("F22").Value
        
    End With
    
    ' MsgBox "Dados salvos com sucesso!", vbInformation
    
    OptimizeCodeExecution False
    
End Sub

Sub RetrieveDataFromName(Optional ShowOnMacroList As Boolean = False)
    
    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchName As String
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Get the ID to search from ComboBox
    If wsForm.OLEObjects("ComboBoxName").Object.Value <> "" Then
        searchName = wsForm.OLEObjects("ComboBoxName").Object.Value
    Else
        'ClearForm
        Exit Sub
    End If
    
    ' Search for the matching row
    Set foundRow = Nothing
    For Each cell In dadosTable.ListColumns(2).DataBodyRange
        If cell.Value & " - " & cell.Offset(0, 1).Value & " - " & cell.Offset(0, 15).Value = searchName Then
            Set foundRow = cell.Offset(0, -1)
            Exit For
        End If
    Next cell
    
    ' If Name is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "Nenhuma obra encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxID").Object.Value = foundRow.Value
    
        ' Read column B values
        .Range("B6").Value = foundRow.Offset(0, 1).Value
        .Range("B10").Value = foundRow.Offset(0, 2).Value
        .Range("B14").Value = foundRow.Offset(0, 3).Value
        .Range("B18").Value = foundRow.Offset(0, 4).Value
        .Range("B22").Value = foundRow.Offset(0, 5).Value
        .Range("B28").Value = foundRow.Offset(0, 6).Value
        .Range("B32").Value = foundRow.Offset(0, 7).Value
        .Range("B36").Value = foundRow.Offset(0, 8).Value
        .Range("B40").Value = foundRow.Offset(0, 9).Value
        .Range("B44").Value = foundRow.Offset(0, 10).Value
        .Range("B48").Value = foundRow.Offset(0, 11).Value
        .Range("B52").Value = foundRow.Offset(0, 12).Value
        .Range("B56").Value = foundRow.Offset(0, 14).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Offset(0, 16).Value
        .Range("D10").Value = foundRow.Offset(0, 17).Value
        .Range("D14").Value = foundRow.Offset(0, 18).Value
        .Range("D18").Value = foundRow.Offset(0, 19).Value
        .Range("D22").Value = foundRow.Offset(0, 20).Value
        .Range("D26").Value = foundRow.Offset(0, 21).Value
        .Range("D30").Value = foundRow.Offset(0, 22).Value
        .Range("D34").Value = foundRow.Offset(0, 23).Value
        .Range("D38").Value = foundRow.Offset(0, 24).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Offset(0, 25).Value
        .Range("F10").Value = foundRow.Offset(0, 26).Value
        .Range("F14").Value = foundRow.Offset(0, 27).Value
        .Range("F18").Value = foundRow.Offset(0, 28).Value
        .Range("F22").Value = foundRow.Offset(0, 30).Value
    End With
    
    OptimizeCodeExecution False

End Sub

Sub RetrieveDataFromID(Optional ShowOnMacroList As Boolean = False)

    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchID As Double
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    wsForm.OLEObjects("ComboBoxName").Top = wsForm.OLEObjects("ComboBoxID").Top + 38
    wsForm.OLEObjects("ComboBoxName").Left = wsForm.OLEObjects("ComboBoxID").Left
    
    ' Get the ID to search from ComboBox
    If wsForm.OLEObjects("ComboBoxID").Object.Value <> "" Then
        searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    Else
        'ClearForm
        Exit Sub
    End If
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(1).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID não encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Offset(0, 1).Value & " - " & foundRow.Offset(0, 2).Value & " - " & foundRow.Offset(0, 16).Value
        
        ' Read column B values
        .Range("B6").Value = foundRow.Offset(0, 1).Value
        .Range("B10").Value = foundRow.Offset(0, 2).Value
        .Range("B14").Value = foundRow.Offset(0, 3).Value
        .Range("B18").Value = foundRow.Offset(0, 4).Value
        .Range("B22").Value = foundRow.Offset(0, 5).Value
        .Range("B28").Value = foundRow.Offset(0, 6).Value
        .Range("B32").Value = foundRow.Offset(0, 7).Value
        .Range("B36").Value = foundRow.Offset(0, 8).Value
        .Range("B40").Value = foundRow.Offset(0, 9).Value
        .Range("B44").Value = foundRow.Offset(0, 10).Value
        .Range("B48").Value = foundRow.Offset(0, 11).Value
        .Range("B52").Value = foundRow.Offset(0, 12).Value
        .Range("B56").Value = foundRow.Offset(0, 14).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Offset(0, 16).Value
        .Range("D10").Value = foundRow.Offset(0, 17).Value
        .Range("D14").Value = foundRow.Offset(0, 18).Value
        .Range("D18").Value = foundRow.Offset(0, 19).Value
        .Range("D22").Value = foundRow.Offset(0, 20).Value
        .Range("D26").Value = foundRow.Offset(0, 21).Value
        .Range("D30").Value = foundRow.Offset(0, 22).Value
        .Range("D34").Value = foundRow.Offset(0, 23).Value
        .Range("D38").Value = foundRow.Offset(0, 24).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Offset(0, 25).Value
        .Range("F10").Value = foundRow.Offset(0, 26).Value
        .Range("F14").Value = foundRow.Offset(0, 27).Value
        .Range("F18").Value = foundRow.Offset(0, 28).Value
        .Range("F22").Value = foundRow.Offset(0, 30).Value
    End With
    
    OptimizeCodeExecution False

End Sub

Sub EnviarParaAprovação(Optional ShowOnMacroList As Boolean = False)

    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    
    Dim OutApp As Object
    Dim OutMail As Object
    
    '--- Variables for email content
    Dim HTMLbody As String
    Dim greeting As String
    Dim strSignature As String
    Dim faseObra As String
    
    '--- Create Outlook instance and a new mail item
    On Error Resume Next
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    On Error GoTo 0
    
    If OutApp Is Nothing Then
        MsgBox "O Outlook não está instalado nesse computador.", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' não encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Get the ID to search from ComboBox
    searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' Stop if data not saved
    If searchID = "" Then
        MsgBox "Desculpe, salve os dados antes de gerar o e-mail", vbInformation, "Atenção"
        Exit Sub
    End If
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(1).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID não encontrado!", vbExclamation
        Exit Sub
    End If
    
    If foundRow.Offset(0, 29).Value <> "" Then
        userResponse = MsgBox("O e-mail de aprovação para esses dados já foi enviado em " & foundRow.Offset(0, 29).Value & ". Deseja enviar novamente?", vbYesNo)
        If userResponse = vbNo Then
            MsgBox "Envio de e-mail cancelado!", vbInformation
            Exit Sub
        End If
    End If
    
    ' Decide between Bom dia or Boa tarde
    If Hour(Now) < 12 Then
        greeting = "Bom dia"
    Else
        greeting = "Boa tarde"
    End If
    
    ' Get user signature
    With OutMail
        .Display ' This opens the email and loads the default signature
        strSignature = .HTMLbody ' Capture the signature
    End With
    
    HTMLbody = ""
    HTMLbody = HTMLbody & "<p>" & greeting & ", Bruna</p>"
    HTMLbody = HTMLbody & "<p>Gentileza dar continuidade na Suplementação conforme abaixo:</p>"
    'HTMLbody = HTMLbody & "<p>Solicito sua confirmação (“De acordo”) quanto aos valores abai xo, para que possamos dar continuidade à contratação da " & _
        foundRow.Offset(0, 22).Value & " para o serviço descrito a seguir: " & foundRow.Offset(0, 15).Value & " da " & foundRow.Offset(0, 1).Value & _
        " no valor de " & Format(foundRow.Offset(0, 7).Value, "R$ #,##0.00") & ". Todos os valores apresentados abaixo foram analisados pela equipe de Implantação/Suprimentos e considerado procedentes." & "</p>"
    
    ' Start the table
    HTMLbody = HTMLbody & "<table border='1' style='border-collapse: collapse; font-size: 10pt;'>"
    
    ' Title row
    HTMLbody = HTMLbody & "<tr style='background-color:#d9d9d9;'>"
    HTMLbody = HTMLbody & "<td colspan='2'><b>Suplementação de Custos" & " - " & foundRow.Offset(0, 1).Value & " - " & foundRow.Offset(0, 2).Value & "</b></td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 1) CUSTO DA SUPLEMENTAÇÃO
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo da suplementação</b></td>"
    ' Example: reading from the "Dados" sheet. Adjust the range as needed.
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 7).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 2) Inserido no DR/tarefa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Inserido no DR/tarefa</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 6).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 3) PEP
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>PEP</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 5).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 4) CUSTO COT DO DR
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo COT do DR</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 8).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 5) CUSTO PLANEJADO ANTES DA SUPLEMENTAÇÃO
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo planejado antes da suplementação (versão Z)</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 9).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
      
    ' 6) CUSTO PLANEJADO APÓS A SUPLEMENTAÇÃO
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo planejado após suplementação</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 10).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 7) Resultado planejado atual antes da suplementação
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Resultado planejado atual antes da suplementação (%)</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 11).Value, "##.00%") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 8) Resultado planejado atual após a suplementação
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Resultado planejado atual após a suplementação (%)</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 12).Value, "##.00%") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 9) Saldo da provisão
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Saldo atual da provisão de riscos no momento da análise:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 15).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 10) Justificativa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Justificativa:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 17).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 11) Outros riscos já mapeados
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Outros riscos já mapeados</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 24).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 12) Estágio da Obra
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Estágio da obra</b></td>"
    
    If IsNumeric(foundRow.Offset(0, 18).Value) Then
        If foundRow.Offset(0, 18).Value < 0.4 Then
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 18).Value, "##.00%") & " (Fase Inicial)" & "</td>"
        ElseIf foundRow.Offset(0, 18).Value < 0.8 Then
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 18).Value, "##.00%") & " (Fase Intermediária)" & "</td>"
        Else
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Offset(0, 18).Value, "##.00%") & " (Fase Final)" & "</td>"
        End If
    Else
        HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 18).Value & "</td>"
    End If

    HTMLbody = HTMLbody & "</tr>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 13) Ação necessária
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Ação necessária</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Offset(0, 20).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' Close the table
    HTMLbody = HTMLbody & "</table>"
    
    '-------------------------------------------------------------------------
    ' Configure and send the email
    '-------------------------------------------------------------------------
    With OutMail
        .To = "emoretti@weg.net"
        .CC = "matheusp@weg.net"
        .BCC = ""
        .Subject = "Aprovação de Custos - Suplementação de Custos - " & foundRow.Offset(0, 1).Value & " - " & foundRow.Offset(0, 2).Value
        .HTMLbody = HTMLbody & strSignature
        .Display   'Use .Display to just open the email draft
        ' .Send       'Use .Send to send immediately
    End With
    
    '--- Cleanup
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    foundRow.Offset(0, 29).Value = Date
    
    MsgBox "Email """ & "Aprovação de Custos - Suplementação de Custos - " & foundRow.Offset(0, 1).Value & " - " & foundRow.Offset(0, 2).Value & """ enviado com sucesso!", vbInformation
    
    OptimizeCodeExecution False
    
End Sub

Sub ClearForm(Optional ShowOnMacroList As Boolean = False)
    
    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    
    If wsForm.OLEObjects("ComboBoxID").Object.Value = "" Then
        If MsgBox("Esses dados não foram salvos. Deseja limpá-los mesmo assim?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        .OLEObjects("ComboBoxID").Object.Value = ""
        .OLEObjects("ComboBoxName").Object.Value = ""
        .OLEObjects("ComboBoxName").Width = 123
        
        ' Read column B values
        .Range("B6").Value = ""
        .Range("B10").Value = ""
        .Range("B14").Value = ""
        .Range("B18").Value = ""
        .Range("B22").Value = ""
        .Range("B28").Value = ""
        .Range("B32").Value = ""
        .Range("B36").Value = ""
        .Range("B40").Value = ""
        .Range("B44").Value = ""
        .Range("B48").Value = ""
        .Range("B52").Value = ""
        .Range("B56").Value = ""
        
        ' Read column D values
        .Range("D6").Value = ""
        .Range("D10").Value = ""
        .Range("D14").Value = ""
        .Range("D18").Value = ""
        .Range("D22").Value = ""
        .Range("D26").Value = ""
        .Range("D30").Value = ""
        .Range("D34").Value = ""
        .Range("D38").Value = ""
        
        ' Read column F values
        .Range("F6").Value = ""
        .Range("F10").Value = ""
        .Range("F14").Value = ""
        .Range("F18").Value = ""
        .Range("F22").Value = ""
    End With
    
    OptimizeCodeExecution False
    
End Sub

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
