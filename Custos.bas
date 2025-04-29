Attribute VB_Name = "Custos"
' ----- Version -----
'        1.3.0
' -------------------

Sub SaveData(Optional ShowOnMacroList As Boolean = False)

    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim tblRow As ListRow
    Dim newID As String
    Dim userResponse As VbMsgBoxResult
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' n�o encontrada!", vbExclamation
        Exit Sub
    End If
    
    newID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' If ComboBoxID is not empty, prompt the user
    If Trim(newID) <> "" Then
        userResponse = MsgBox("Esse aditivo j� foi cadastrado. Deseja sobrescrever?", vbYesNoCancel + vbQuestion, "Confirma��o")

        Select Case userResponse
            Case vbYes
                newID = Val(newID) ' Use ComboBoxID.Value as new ID
                ' Search for the ID in the first column of the table
                Set tblRow = dadosTable.ListRows(dadosTable.ListColumns(colMap("ID")).DataBodyRange.Find(What:=newID, LookAt:=xlWhole).Row - dadosTable.DataBodyRange.Row + 1)
            Case vbNo
                ' Proceed with generating new ID
                newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(colMap("ID")).DataBodyRange) + 1
                wsForm.OLEObjects("ComboBoxID").Object.Value = newID
                ' Add a new row to the table
                Set tblRow = dadosTable.ListRows.Add
            Case vbCancel
                Exit Sub ' Exit without saving
        End Select
    Else
        newID = Application.WorksheetFunction.Max(dadosTable.ListColumns(colMap("ID")).DataBodyRange) + 1
        
        wsForm.OLEObjects("ComboBoxID").Object.Value = newID
        
        wsForm.OLEObjects("ComboBoxName").Object.Value = wsForm.Range("B6").Value & " - " & wsForm.Range("B10").Value & " - " & wsForm.Range("D6").Value
        
        ' Add a new row to the table
        Set tblRow = dadosTable.ListRows.Add
    End If
    
    ' Assign values to the new row
    With tblRow.Range
        ' Set new ID
        .Cells(1, colMap("ID")).Value = newID ' First column value
        
        ' Read column B values
        .Cells(1, colMap("Obra")).Value = wsForm.Range("B6").Value
        .Cells(1, colMap("Cliente")).Value = wsForm.Range("B10").Value
        .Cells(1, colMap("Tipo")).Value = wsForm.Range("B14").Value
        .Cells(1, colMap("PM")).Value = wsForm.Range("B18").Value
        .Cells(1, colMap("PEP")).Value = wsForm.Range("B22").Value
        .Cells(1, colMap("DR")).Value = wsForm.Range("B28").Value
        .Cells(1, colMap("Suplementa��o")).Value = wsForm.Range("B32").Value
        .Cells(1, colMap("COT")).Value = wsForm.Range("B36").Value
        .Cells(1, colMap("Custo Antes")).Value = wsForm.Range("B40").Value
        .Cells(1, colMap("Custo Depois")).Value = wsForm.Range("B44").Value
        .Cells(1, colMap("Resultado Antes")).Value = wsForm.Range("B48").Value
        .Cells(1, colMap("Resultado Depois")).Value = wsForm.Range("B52").Value
        
        If .Cells(1, colMap("Impacto")).Formula = "" Then
            .Cells(1, colMap("Impacto")).Formula = "=IFERROR([@[Custo da Suplementa��o]]/[@[Custo COT]];"")"
        End If
        If .Cells(1, colMap("Saldo")).Formula = "" Then
            .Cells(1, colMap("Saldo")).Formula = "=[@[Custo Planejado Depois da Suplementa��o]]-[@[Custo COT]]-[@[Custo da Suplementa��o]]"
        End If
        
        .Cells(1, colMap("Provis�o")).Value = wsForm.Range("B56").Value
        
        ' Read column D values
        .Cells(1, colMap("Descri��o")).Value = wsForm.Range("D6").Value
        .Cells(1, colMap("Justificativa")).Value = wsForm.Range("D10").Value
        
        If wsForm.Range("D14").Value < 0.4 Then
            .Cells(1, colMap("Est�gio")).Value = Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Inicial)"
        ElseIf wsForm.Range("D14").Value < 0.8 Then
            .Cells(1, colMap("Est�gio")).Value = Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Intermedi�ria)"
        Else
            .Cells(1, colMap("Est�gio")).Value = Format(wsForm.Range("D14").Value, "##.00%") & " (Fase Final)"
        End If

        .Cells(1, colMap("Fator")).Value = wsForm.Range("D18").Value
        .Cells(1, colMap("Detalhamento")).Value = wsForm.Range("D22").Value
        .Cells(1, colMap("Repasse")).Value = wsForm.Range("D26").Value
        .Cells(1, colMap("Justificativa Repasse")).Value = wsForm.Range("D30").Value
        .Cells(1, colMap("Prestador")).Value = wsForm.Range("D34").Value
        .Cells(1, colMap("Riscos")).Value = wsForm.Range("D38").Value
    
        ' Read column F values
        .Cells(1, colMap("Status")).Value = wsForm.Range("F6").Value
        .Cells(1, colMap("RFP")).Value = wsForm.Range("F10").Value
        .Cells(1, colMap("Suprimentos")).Value = wsForm.Range("F14").Value
        .Cells(1, colMap("Pedido")).Value = wsForm.Range("F18").Value
        .Cells(1, colMap("Data")).Value = "" ' Erase date if information is overwriten to allow resend the e-mail
        .Cells(1, colMap("Observa��es")).Value = wsForm.Range("F22").Value
        
    End With
    
    ' MsgBox "Dados salvos com sucesso!", vbInformation
    
    OptimizeCodeExecution False
    
End Sub

Sub RetrieveDataFromName(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchName As String
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' n�o encontrada!", vbExclamation
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
    For Each cell In dadosTable.ListColumns(colMap("ID")).DataBodyRange
        If cell.Value & " - " & cell.Cells(1, colMap("Cliente")) & " - " & cell.Cells(1, colMap("Obra")) & " - " & cell.Cells(1, colMap("Descri��o")) = searchName Then
            Set foundRow = cell
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
        wsForm.OLEObjects("ComboBoxID").Object.Value = foundRow.Cells(1, colMap("ID")).Value
    
        ' Read column B values
        .Range("B6").Value = foundRow.Cells(1, colMap("Obra")).Value
        .Range("B10").Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range("B14").Value = foundRow.Cells(1, colMap("Tipo")).Value
        .Range("B18").Value = foundRow.Cells(1, colMap("PM")).Value
        .Range("B22").Value = foundRow.Cells(1, colMap("PEP")).Value
        .Range("B28").Value = foundRow.Cells(1, colMap("DR")).Value
        .Range("B32").Value = foundRow.Cells(1, colMap("Suplementa��o")).Value
        .Range("B36").Value = foundRow.Cells(1, colMap("COT")).Value
        .Range("B40").Value = foundRow.Cells(1, colMap("Custo Antes")).Value
        .Range("B44").Value = foundRow.Cells(1, colMap("Custo Depois")).Value
        .Range("B48").Value = foundRow.Cells(1, colMap("Resultado Antes")).Value
        .Range("B52").Value = foundRow.Cells(1, colMap("Resultado Depois")).Value
        .Range("B56").Value = foundRow.Cells(1, colMap("Provis�o")).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Cells(1, colMap("Descri��o")).Value
        .Range("D10").Value = foundRow.Cells(1, colMap("Justificativa")).Value
        .Range("D14").Value = foundRow.Cells(1, colMap("Est�gio")).Value
        .Range("D18").Value = foundRow.Cells(1, colMap("Fator")).Value
        .Range("D22").Value = foundRow.Cells(1, colMap("Detalhamento")).Value
        .Range("D26").Value = foundRow.Cells(1, colMap("Repasse")).Value
        .Range("D30").Value = foundRow.Cells(1, colMap("Justificativa Repasse")).Value
        .Range("D34").Value = foundRow.Cells(1, colMap("Prestador")).Value
        .Range("D38").Value = foundRow.Cells(1, colMap("Riscos")).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Cells(1, colMap("Status")).Value
        .Range("F10").Value = foundRow.Cells(1, colMap("RFP")).Value
        .Range("F14").Value = foundRow.Cells(1, colMap("Suprimentos")).Value
        .Range("F18").Value = foundRow.Cells(1, colMap("Pedido")).Value
        .Range("F22").Value = foundRow.Cells(1, colMap("Observa��es")).Value
    End With
    
    OptimizeCodeExecution False

End Sub

Sub RetrieveDataFromID(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet, wsDados As Worksheet
    Dim dadosTable As ListObject
    Dim foundRow As Range
    Dim searchID As Double
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' n�o encontrada!", vbExclamation
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
    Set foundRow = dadosTable.ListColumns(colMap("ID")).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID n�o encontrado!", vbExclamation
        Exit Sub
    End If
    
    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Value & " - " & foundRow.Cells(1, colMap("Cliente")) & " - " & foundRow.Cells(1, colMap("Obra")) & " - " & foundRow.Cells(1, colMap("Descri��o"))
        
        ' Read column B values
        .Range("B6").Value = foundRow.Cells(1, colMap("Obra")).Value
        .Range("B10").Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range("B14").Value = foundRow.Cells(1, colMap("Tipo")).Value
        .Range("B18").Value = foundRow.Cells(1, colMap("PM")).Value
        .Range("B22").Value = foundRow.Cells(1, colMap("PEP")).Value
        .Range("B28").Value = foundRow.Cells(1, colMap("DR")).Value
        .Range("B32").Value = foundRow.Cells(1, colMap("Suplementa��o")).Value
        .Range("B36").Value = foundRow.Cells(1, colMap("COT")).Value
        .Range("B40").Value = foundRow.Cells(1, colMap("Custo Antes")).Value
        .Range("B44").Value = foundRow.Cells(1, colMap("Custo Depois")).Value
        .Range("B48").Value = foundRow.Cells(1, colMap("Resultado Antes")).Value
        .Range("B52").Value = foundRow.Cells(1, colMap("Resultado Depois")).Value
        .Range("B56").Value = foundRow.Cells(1, colMap("Provis�o")).Value
        
        ' Read column D values
        .Range("D6").Value = foundRow.Cells(1, colMap("Descri��o")).Value
        .Range("D10").Value = foundRow.Cells(1, colMap("Justificativa")).Value
        .Range("D14").Value = foundRow.Cells(1, colMap("Est�gio")).Value
        .Range("D18").Value = foundRow.Cells(1, colMap("Fator")).Value
        .Range("D22").Value = foundRow.Cells(1, colMap("Detalhamento")).Value
        .Range("D26").Value = foundRow.Cells(1, colMap("Repasse")).Value
        .Range("D30").Value = foundRow.Cells(1, colMap("Justificativa Repasse")).Value
        .Range("D34").Value = foundRow.Cells(1, colMap("Prestador")).Value
        .Range("D38").Value = foundRow.Cells(1, colMap("Riscos")).Value
        
        ' Read column F values
        .Range("F6").Value = foundRow.Cells(1, colMap("Status")).Value
        .Range("F10").Value = foundRow.Cells(1, colMap("RFP")).Value
        .Range("F14").Value = foundRow.Cells(1, colMap("Suprimentos")).Value
        .Range("F18").Value = foundRow.Cells(1, colMap("Pedido")).Value
        .Range("F22").Value = foundRow.Cells(1, colMap("Observa��es")).Value
    End With
    
    OptimizeCodeExecution False

End Sub

Sub EnviarParaAprova��o(Optional ShowOnMacroList As Boolean = False)
    
    Dim colMap As Object
    Set colMap = GetColumnHeadersMapping()
    
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
        MsgBox "O Outlook n�o est� instalado nesse computador.", vbExclamation
        Exit Sub
    End If
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    Set wsDados = ThisWorkbook.Sheets("Dados")
    
    ' Check if table "Dados" exists
    On Error Resume Next
    Set dadosTable = wsDados.ListObjects("Dados")
    On Error GoTo 0
    
    ' If the table doesn't exist, exit sub
    If dadosTable Is Nothing Then
        MsgBox "Tabela 'Dados' n�o encontrada!", vbExclamation
        Exit Sub
    End If
    
    ' Get the ID to search from ComboBox
    searchID = wsForm.OLEObjects("ComboBoxID").Object.Value
    
    ' Stop if data not saved
    If searchID = "" Then
        MsgBox "Desculpe, salve os dados antes de gerar o e-mail", vbInformation, "Aten��o"
        Exit Sub
    End If
    
    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(colMap("ID")).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0
    
    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID n�o encontrado!", vbExclamation
        Exit Sub
    End If
    
    If foundRow.Cells(1, 29).Value <> "" Then
        userResponse = MsgBox("O e-mail de aprova��o para esses dados j� foi enviado em " & foundRow.Cells(1, 29).Value & ". Deseja enviar novamente?", vbYesNo)
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
    HTMLbody = HTMLbody & "<p>Gentileza dar continuidade na Suplementa��o conforme abaixo:</p>"
    'HTMLbody = HTMLbody & "<p>Solicito sua confirma��o (�De acordo�) quanto aos valores abai xo, para que possamos dar continuidade � contrata��o da " & _
        foundRow.Cells(1, 22).Value & " para o servi�o descrito a seguir: " & foundRow.Cells(1, 15).Value & " da " & foundRow.Cells(1, 1).Value & _
        " no valor de " & Format(foundRow.Cells(1, 7).Value, "R$ #,##0.00") & ". Todos os valores apresentados abaixo foram analisados pela equipe de Implanta��o/Suprimentos e considerado procedentes." & "</p>"
    
    ' Start the table
    HTMLbody = HTMLbody & "<table border='1' style='border-collapse: collapse; font-size: 10pt;'>"
    
    ' Title row
    HTMLbody = HTMLbody & "<tr style='background-color:#d9d9d9;'>"
    HTMLbody = HTMLbody & "<td colspan='2'><b>Suplementa��o de Custos" & " - " & foundRow.Cells(1, 1).Value & " - " & foundRow.Cells(1, 2).Value & "</b></td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 1) CUSTO DA SUPLEMENTA��O
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo da suplementa��o</b></td>"
    ' Example: reading from the "Dados" sheet. Adjust the range as needed.
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 7).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 2) Inserido no DR/tarefa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Inserido no DR/tarefa</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, 6).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 3) PEP
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>PEP</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, 5).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 4) CUSTO COT DO DR
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo COT do DR</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 8).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 5) CUSTO PLANEJADO ANTES DA SUPLEMENTA��O
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo planejado antes da suplementa��o (vers�o Z)</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 9).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
      
    ' 6) CUSTO PLANEJADO AP�S A SUPLEMENTA��O
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo planejado ap�s suplementa��o</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 10).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 7) Resultado planejado atual antes da suplementa��o
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Resultado planejado atual antes da suplementa��o (%)</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 11).Value, "##.00%") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 8) Resultado planejado atual ap�s a suplementa��o
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Resultado planejado atual ap�s a suplementa��o (%)</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 12).Value, "##.00%") & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 9) Saldo da provis�o
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Saldo atual da provis�o de riscos no momento da an�lise:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, 15).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 10) Justificativa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Justificativa:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, 17).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 11) Outros riscos j� mapeados
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Outros riscos j� mapeados</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, 24).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 12) Est�gio da Obra
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Est�gio da obra</b></td>"
    
    If IsNumeric(foundRow.Cells(1, 18).Value) Then
        If foundRow.Cells(1, 18).Value < 0.4 Then
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 18).Value, "##.00%") & " (Fase Inicial)" & "</td>"
        ElseIf foundRow.Cells(1, 18).Value < 0.8 Then
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 18).Value, "##.00%") & " (Fase Intermedi�ria)" & "</td>"
        Else
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, 18).Value, "##.00%") & " (Fase Final)" & "</td>"
        End If
    Else
        HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, 18).Value & "</td>"
    End If

    HTMLbody = HTMLbody & "</tr>"
    HTMLbody = HTMLbody & "</tr>"
    
    ' 13) A��o necess�ria
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>A��o necess�ria</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, 20).Value & "</td>"
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
        .Subject = "Aprova��o de Custos - Suplementa��o de Custos - " & foundRow.Cells(1, 1).Value & " - " & foundRow.Cells(1, 2).Value
        .HTMLbody = HTMLbody & strSignature
        .Display   'Use .Display to just open the email draft
        ' .Send       'Use .Send to send immediately
    End With
    
    '--- Cleanup
    Set OutMail = Nothing
    Set OutApp = Nothing
    
    foundRow.Cells(1, 29).Value = Date
    
    MsgBox "Email """ & "Aprova��o de Custos - Suplementa��o de Custos - " & foundRow.Cells(1, 1).Value & " - " & foundRow.Cells(1, 2).Value & """ enviado com sucesso!", vbInformation
    
    OptimizeCodeExecution False
    
End Sub

Sub ClearForm(Optional ShowOnMacroList As Boolean = False)

    OptimizeCodeExecution True
    
    Dim wsForm As Worksheet
    
    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formul�rio")
    
    If wsForm.OLEObjects("ComboBoxID").Object.Value = "" Then
        If MsgBox("Esses dados n�o foram salvos. Deseja limp�-los mesmo assim?", vbYesNo) = vbNo Then
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

Public Function GetColumnHeadersMapping() As Object
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")
    
    ' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    headers.Add "ID", 1
    headers.Add "Obra", 2
    headers.Add "Cliente", 3
    headers.Add "Tipo", 4
    headers.Add "PM", 5
    headers.Add "PEP", 6
    headers.Add "DR", 7
    headers.Add "Suplementa��o", 8
    headers.Add "COT", 9
    headers.Add "Custo Antes", 10
    headers.Add "Custo Depois", 11
    headers.Add "Resultado Antes", 12
    headers.Add "Resultado Depois", 13
    headers.Add "Impacto", 14
    headers.Add "Saldo", 15
    headers.Add "Provis�o", 16
    headers.Add "Descri��o", 17
    headers.Add "Justificativa", 18
    headers.Add "Est�gio", 19
    headers.Add "Fator", 20
    headers.Add "Detalhamento", 21
    headers.Add "Repasse", 22
    headers.Add "Justificativa Repasse", 23
    headers.Add "Prestador", 24
    headers.Add "Riscos", 25
    headers.Add "Status", 26
    headers.Add "RFP", 27
    headers.Add "Suprimentos", 28
    headers.Add "Pedido", 29
    headers.Add "Data", 30
    headers.Add "Observa��es", 31
    headers.Add "Atividade", 32
    headers.Add "ATV", 33
    
    Set GetColumnHeadersMapping = headers
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
