Attribute VB_Name = "Custos"
' ----- Version -----
'        1.4.0
' -------------------

Sub SaveData(Optional ShowOnMacroList As Boolean = False)

    Dim colMap As Object, formMap As Object
    Set colMap = GetColumnHeadersMapping()
    Set formMap = GetFormFieldsMapping()

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
        MsgBox "Tabela 'Dados' nao encontrada!", vbExclamation
        Exit Sub
    End If

    newID = wsForm.OLEObjects("ComboBoxID").Object.Value

    ' If ComboBoxID is not empty, prompt the user
    If Trim(newID) <> "" Then
        userResponse = MsgBox("Esse aditivo ja foi cadastrado. Deseja sobrescrevera", vbYesNoCancel + vbQuestion, "Confirmaçao")

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

        wsForm.OLEObjects("ComboBoxName").Object.Value = wsForm.Range(formMap("Obra")).Value & " - " & wsForm.Range(formMap("Cliente")).Value & " - " & wsForm.Range(formMap("Descriçao")).Value

        ' Add a new row to the table
        Set tblRow = dadosTable.ListRows.Add
    End If

    ' Assign values to the new row
    With tblRow.Range
        ' Set new ID
        .Cells(1, colMap("ID")).Value = newID ' First column value

        ' Read column B values
        .Cells(1, colMap("Obra")).Value = wsForm.Range(formMap("Obra")).Value
        .Cells(1, colMap("Cliente")).Value = wsForm.Range(formMap("Cliente")).Value
        .Cells(1, colMap("Tipo")).Value = wsForm.Range(formMap("Tipo")).Value
        .Cells(1, colMap("PM")).Value = wsForm.Range(formMap("PM")).Value
        .Cells(1, colMap("PEP")).Value = wsForm.Range(formMap("PEP")).Value
        '.Cells(1, colMap("DR")).Value = wsForm.Range(formMap("DR")).Value
        .Cells(1, colMap("Suplementacao")).Value = wsForm.Range(formMap("Suplementacao")).Value
        .Cells(1, colMap("COT")).Value = wsForm.Range(formMap("COT")).Value
        .Cells(1, colMap("Custo Antes")).Value = wsForm.Range(formMap("Custo Antes")).Value
        .Cells(1, colMap("Custo Depois")).Value = wsForm.Range(formMap("Custo Depois")).Value
        .Cells(1, colMap("Resultado Inicial")).Value = wsForm.Range(formMap("Resultado Inicial")).Value
        .Cells(1, colMap("Resultado Antes")).Value = wsForm.Range(formMap("Resultado Antes")).Value
        .Cells(1, colMap("Resultado Depois")).Value = wsForm.Range(formMap("Resultado Depois")).Value

        If .Cells(1, colMap("Impacto")).Formula = "" Then
            .Cells(1, colMap("Impacto")).Formula = "=IFERROR([@[Custo da Suplementaçao]]/[@[Custo COT]];"")"
        End If
        If .Cells(1, colMap("Saldo")).Formula = "" Then
            .Cells(1, colMap("Saldo")).Formula = "=[@[Custo Planejado Depois da Suplementaçao]]-[@[Custo COT]]-[@[Custo da Suplementaçao]]"
        End If

        .Cells(1, colMap("Provisao")).Value = wsForm.Range(formMap("Provisao")).Value

        ' Read column D values
        .Cells(1, colMap("Descricao")).Value = wsForm.Range(formMap("Descricao")).Value
        .Cells(1, colMap("Justificativa")).Value = wsForm.Range(formMap("Justificativa")).Value

        If Val(wsForm.Range(formMap("Estagio")).Value) / 100 < 0.4 Then
            .Cells(1, colMap("Estagio")).Value = Format(Val(wsForm.Range(formMap("Estagio")).Value) / 100, "##.00%") & " (Fase Inicial)"
        ElseIf Val(wsForm.Range(formMap("Estagio")).Value) / 100 < 0.8 Then
            .Cells(1, colMap("Estagio")).Value = Format(Val(wsForm.Range(formMap("Estagio")).Value) / 100, "##.00%") & " (Fase Intermediaria)"
        Else
            .Cells(1, colMap("Estagio")).Value = Format(Val(wsForm.Range(formMap("Estagio")).Value) / 100, "##.00%") & " (Fase Final)"
        End If

        .Cells(1, colMap("Fator")).Value = wsForm.Range(formMap("Fator")).Value
        .Cells(1, colMap("Detalhamento")).Value = wsForm.Range(formMap("Detalhamento")).Value
        .Cells(1, colMap("Repasse")).Value = wsForm.Range(formMap("Repasse")).Value
        .Cells(1, colMap("Justificativa Repasse")).Value = wsForm.Range(formMap("Justificativa Repasse")).Value
        .Cells(1, colMap("Prestador")).Value = wsForm.Range(formMap("Prestador")).Value
        .Cells(1, colMap("Riscos")).Value = wsForm.Range(formMap("Riscos")).Value

        ' Read column F values
        .Cells(1, colMap("Status")).Value = wsForm.Range(formMap("Status")).Value
        .Cells(1, colMap("RFP")).Value = wsForm.Range(formMap("RFP")).Value
        .Cells(1, colMap("Suprimentos")).Value = wsForm.Range(formMap("Suprimentos")).Value
        .Cells(1, colMap("Pedido")).Value = wsForm.Range(formMap("Pedido")).Value
        .Cells(1, colMap("Data")).Value = "" ' Erase date if information is overwriten to allow resend the e-mail
        .Cells(1, colMap("observacoes")).Value = wsForm.Range(formMap("observacoes")).Value

    End With

    ' MsgBox "Dados salvos com sucesso!", vbInformation

    OptimizeCodeExecution False

End Sub

Sub RetrieveDataFromName(Optional ShowOnMacroList As Boolean = False)

    Dim colMap As Object, formMap As Object
    Set colMap = GetColumnHeadersMapping()
    Set formMap = GetFormFieldsMapping()

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
        MsgBox "Tabela 'Dados' nao encontrada!", vbExclamation
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
        If cell.Value & " - " & cell.Cells(1, colMap("Cliente")) & " - " & cell.Cells(1, colMap("Obra")) & " - " & cell.Cells(1, colMap("Descricao")) = searchName Then
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
        .Range(formMap("Obra")).Value = foundRow.Cells(1, colMap("Obra")).Value
        .Range(formMap("Cliente")).Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range(formMap("Tipo")).Value = foundRow.Cells(1, colMap("Tipo")).Value
        .Range(formMap("PM")).Value = foundRow.Cells(1, colMap("PM")).Value
        .Range(formMap("PEP")).Value = foundRow.Cells(1, colMap("PEP")).Value
        '.Range(formMap("DR")).Value = foundRow.Cells(1, colMap("DR")).Value
        .Range(formMap("Suplementacao")).Value = foundRow.Cells(1, colMap("Suplementacao")).Value
        .Range(formMap("COT")).Value = foundRow.Cells(1, colMap("COT")).Value
        .Range(formMap("Custo Antes")).Value = foundRow.Cells(1, colMap("Custo Antes")).Value
        .Range(formMap("Custo Depois")).Value = foundRow.Cells(1, colMap("Custo Depois")).Value
        .Range(formMap("Resultado Inicial")).Value = foundRow.Cells(1, colMap("Resultado Inicial")).Value
        .Range(formMap("Resultado Antes")).Value = foundRow.Cells(1, colMap("Resultado Antes")).Value
        .Range(formMap("Resultado Depois")).Value = foundRow.Cells(1, colMap("Resultado Depois")).Value
        .Range(formMap("Provisao")).Value = foundRow.Cells(1, colMap("Provisao")).Value

        ' Read column D values
        .Range(formMap("Descricao")).Value = foundRow.Cells(1, colMap("Descricao")).Value
        .Range(formMap("Justificativa")).Value = foundRow.Cells(1, colMap("Justificativa")).Value
        .Range(formMap("Estagio")).Value = foundRow.Cells(1, colMap("Estagio")).Value
        .Range(formMap("Fator")).Value = foundRow.Cells(1, colMap("Fator")).Value
        .Range(formMap("Detalhamento")).Value = foundRow.Cells(1, colMap("Detalhamento")).Value
        .Range(formMap("Repasse")).Value = foundRow.Cells(1, colMap("Repasse")).Value
        .Range(formMap("Justificativa Repasse")).Value = foundRow.Cells(1, colMap("Justificativa Repasse")).Value
        .Range(formMap("Prestador")).Value = foundRow.Cells(1, colMap("Prestador")).Value
        .Range(formMap("Riscos")).Value = foundRow.Cells(1, colMap("Riscos")).Value

        ' Read column F values
        .Range(formMap("Status")).Value = foundRow.Cells(1, colMap("Status")).Value
        .Range(formMap("RFP")).Value = foundRow.Cells(1, colMap("RFP")).Value
        .Range(formMap("Suprimentos")).Value = foundRow.Cells(1, colMap("Suprimentos")).Value
        .Range(formMap("Pedido")).Value = foundRow.Cells(1, colMap("Pedido")).Value
        .Range(formMap("observacoes")).Value = foundRow.Cells(1, colMap("observacoes")).Value
    End With

    OptimizeCodeExecution False

End Sub

Sub RetrieveDataFromID(Optional ShowOnMacroList As Boolean = False)

    Dim colMap As Object, formMap As Object
    Set colMap = GetColumnHeadersMapping()
    Set formMap = GetFormFieldsMapping()

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
        MsgBox "Tabela 'Dados' nao encontrada!", vbExclamation
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
        MsgBox "ID nao encontrado!", vbExclamation
        Exit Sub
    End If

    ' Populate worksheet with retrieved data
    With wsForm
        wsForm.OLEObjects("ComboBoxName").Object.Value = foundRow.Value & " - " & foundRow.Cells(1, colMap("Cliente")) & " - " & foundRow.Cells(1, colMap("Obra")) & " - " & foundRow.Cells(1, colMap("Descricao"))

        ' Read column B values
        .Range(formMap("Obra")).Value = foundRow.Cells(1, colMap("Obra")).Value
        .Range(formMap("Cliente")).Value = foundRow.Cells(1, colMap("Cliente")).Value
        .Range(formMap("Tipo")).Value = foundRow.Cells(1, colMap("Tipo")).Value
        .Range(formMap("PM")).Value = foundRow.Cells(1, colMap("PM")).Value
        .Range(formMap("PEP")).Value = foundRow.Cells(1, colMap("PEP")).Value
        '.Range(formMap("DR")).Value = foundRow.Cells(1, colMap("DR")).Value
        .Range(formMap("Suplementacao")).Value = foundRow.Cells(1, colMap("Suplementacao")).Value
        .Range(formMap("COT")).Value = foundRow.Cells(1, colMap("COT")).Value
        .Range(formMap("Custo Antes")).Value = foundRow.Cells(1, colMap("Custo Antes")).Value
        .Range(formMap("Custo Depois")).Value = foundRow.Cells(1, colMap("Custo Depois")).Value
        .Range(formMap("Resultado Inicial")).Value = foundRow.Cells(1, colMap("Resultado Inicial")).Value
        .Range(formMap("Resultado Antes")).Value = foundRow.Cells(1, colMap("Resultado Antes")).Value
        .Range(formMap("Resultado Depois")).Value = foundRow.Cells(1, colMap("Resultado Depois")).Value
        .Range(formMap("Provisao")).Value = foundRow.Cells(1, colMap("Provisao")).Value

        ' Read column D values
        .Range(formMap("Descricao")).Value = foundRow.Cells(1, colMap("Descricao")).Value
        .Range(formMap("Justificativa")).Value = foundRow.Cells(1, colMap("Justificativa")).Value
        .Range(formMap("Estagio")).Value = foundRow.Cells(1, colMap("Estagio")).Value
        .Range(formMap("Fator")).Value = foundRow.Cells(1, colMap("Fator")).Value
        .Range(formMap("Detalhamento")).Value = foundRow.Cells(1, colMap("Detalhamento")).Value
        .Range(formMap("Repasse")).Value = foundRow.Cells(1, colMap("Repasse")).Value
        .Range(formMap("Justificativa Repasse")).Value = foundRow.Cells(1, colMap("Justificativa Repasse")).Value
        .Range(formMap("Prestador")).Value = foundRow.Cells(1, colMap("Prestador")).Value
        .Range(formMap("Riscos")).Value = foundRow.Cells(1, colMap("Riscos")).Value

        ' Read column F values
        .Range(formMap("Status")).Value = foundRow.Cells(1, colMap("Status")).Value
        .Range(formMap("RFP")).Value = foundRow.Cells(1, colMap("RFP")).Value
        .Range(formMap("Suprimentos")).Value = foundRow.Cells(1, colMap("Suprimentos")).Value
        .Range(formMap("Pedido")).Value = foundRow.Cells(1, colMap("Pedido")).Value
        .Range(formMap("observacoes")).Value = foundRow.Cells(1, colMap("observacoes")).Value
    End With

    OptimizeCodeExecution False

End Sub

Sub EnviarParaAprovacao(Optional ShowOnMacroList As Boolean = False)

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
        MsgBox "O Outlook nao esta instalado nesse computador.", vbExclamation
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
        MsgBox "Tabela 'Dados' nao encontrada!", vbExclamation
        Exit Sub
    End If

    ' Get the ID to search from ComboBox
    searchID = wsForm.OLEObjects("ComboBoxID").Object.Value

    ' Stop if data not saved
    If searchID = "" Then
        MsgBox "Desculpe, salve os dados antes de gerar o e-mail", vbInformation, "Atencao"
        Exit Sub
    End If

    ' Search for the ID in the first column of the table
    Set foundRow = Nothing
    On Error Resume Next
    Set foundRow = dadosTable.ListColumns(colMap("ID")).DataBodyRange.Find(What:=searchID, LookAt:=xlWhole)
    On Error GoTo 0

    ' If ID is not found, exit sub
    If foundRow Is Nothing Then
        MsgBox "ID nao encontrado!", vbExclamation
        Exit Sub
    End If
    
' ===== Validate required fields used in HTMLBody =====
Dim requiredKeys As Variant, k As Variant, idx As Long, v As Variant
Dim missing As String
requiredKeys = Array( _
    "Obra", "Cliente", _
    "Suplementacao", "PEP", _
    "Custo Inicial", "Resultado Inicial", _
    "Custo Antes", "Resultado Antes", _
    "Custo Depois", "Resultado Depois", _
    "Provisao", "Justificativa", "Riscos", _
    "Estagio", "Detalhamento" _
)

For Each k In requiredKeys
    idx = colMap(k)
    v = foundRow.Cells(1, idx).Value
    If IsBlankValue(v) Then
        If Len(missing) > 0 Then missing = missing & ", "
        missing = missing & k
    End If
Next k

If Len(missing) > 0 Then
    MsgBox "Não foi possível gerar o e-mail. Existem campos vazios no formulário.", vbExclamation, "Campos obrigatórios"
    GoTo CleanExit
End If
' ===== End validation =====

    If foundRow.Cells(1, colMap("Data")).Value <> "" Then
        userResponse = MsgBox("O e-mail de aprovacao para esses dados ja foi enviado em " & foundRow.Cells(1, colMap("Data")).Value & ". Deseja enviar novamentea", vbYesNo)
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
    HTMLbody = HTMLbody & "<p>" & greeting & ", Prezado</p>"
    HTMLbody = HTMLbody & "<p>Gentileza dar continuidade na Suplementacao conforme abaixo:</p>"
    'HTMLbody = HTMLbody & "<p>Solicito sua confirmacao (aDe acordoa) quanto aos valores abai xo, para que possamos dar continuidade a contratacao da " & _
        foundRow.Cells(1, 22).Value & " para o serviao descrito a seguir: " & foundRow.Cells(1, 15).Value & " da " & foundRow.Cells(1, 1).Value & _
        " no valor de " & Format(foundRow.Cells(1, 7).Value, "R$ #,##0.00") & ". Todos os valores apresentados abaixo foram analisados pela equipe de Implantacao/Suprimentos e considerado procedentes." & "</p>"

    ' Start the table
    HTMLbody = HTMLbody & "<table border='1' style='border-collapse: collapse; font-size: 10pt;'>"

    ' Title row
    HTMLbody = HTMLbody & "<tr style='background-color:#d9d9d9;'>"
    HTMLbody = HTMLbody & "<td colspan='2'><b>Suplementacao de Custos" & " - " & foundRow.Cells(1, colMap("Obra")).Value & "/" & foundRow.Cells(1, colMap("Cliente")).Value & "</b></td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 1) CUSTO DA SUPLEMENTAcaO
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo desta suplementacao</b></td>"
    ' Example: reading from the "Dados" sheet. Adjust the range as needed.
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Suplementacao")).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 2) Inserido no DR/tarefa
    'HTMLbody = HTMLbody & "<tr>"
    'HTMLbody = HTMLbody & "<td><b>Inserido no DR/tarefa</b></td>"
    'HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("DR")).Value & "</td>"
    'HTMLbody = HTMLbody & "</tr>"

    ' 3) PEP
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>PEP</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("PEP")).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 4) Custo/Margem (%) COT do PEP (inicial)
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo/Margem (%) COT do PEP (inicial)</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Custo Inicial")).Value, "R$ #,##0.00") & " / " & Format(foundRow.Cells(1, colMap("Resultado Inicial")).Value, "##.00%") & "</td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 5) Custo/Margem(%) planejado ata essa suplementacao
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo/Margem(%) planejado apas esta suplementacao</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Custo Antes")).Value, "R$ #,##0.00") & " / " & Format(foundRow.Cells(1, colMap("Resultado Antes")).Value, "##.00%") & "</td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 6) Custo/Margem(%) planejado apas esta suplementacao
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Custo/Margem(%) planejado apas esta suplementacao</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Custo Depois")).Value, "R$ #,##0.00") & " / " & Format(foundRow.Cells(1, colMap("Resultado Depois")).Value, "##.00%") & "</td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 7) Resultado planejado atual antes da suplementacao
    'HTMLbody = HTMLbody & "<tr>"
    'HTMLbody = HTMLbody & "<td><b>Resultado planejado atual antes da suplementacao (%)</b></td>"
    'HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Resultado Antes")).Value, "##.00%") & "</td>"
    'HTMLbody = HTMLbody & "</tr>"

    ' 8) Resultado planejado atual apas a suplementacao
    'HTMLbody = HTMLbody & "<tr>"
    'HTMLbody = HTMLbody & "<td><b>Resultado planejado atual apas a suplementacao (%)</b></td>"
    'HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Resultado Depois")).Value, "##.00%") & "</td>"
    'HTMLbody = HTMLbody & "</tr>"

    ' 9) Saldo da provisao
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Saldo atual da provisao de riscos no momento da analise:</b></td>"
    HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Provisao")).Value, "R$ #,##0.00") & "</td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 10) Justificativa
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Justificativa:</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Justificativa")).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 11) Outros riscos ja mapeados
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Outros riscos ja mapeados</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Riscos")).Value & "</td>"
    HTMLbody = HTMLbody & "</tr>"

    ' 12) Estagio da Obra
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Estagio da obra</b></td>"

    If IsNumeric(foundRow.Cells(1, colMap("Estagio")).Value) Then
        If foundRow.Cells(1, colMap("Estagio")).Value < 0.4 Then
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Estagio")).Value, "##.00%") & " (Fase Inicial)" & "</td>"
        ElseIf foundRow.Cells(1, colMap("Estagio")).Value < 0.8 Then
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Estagio")).Value, "##.00%") & " (Fase Intermediaria)" & "</td>"
        Else
            HTMLbody = HTMLbody & "<td>" & Format(foundRow.Cells(1, colMap("Estagio")).Value, "##.00%") & " (Fase Final)" & "</td>"
        End If
    Else
        HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Estagio")).Value & "</td>"
    End If

    HTMLbody = HTMLbody & "</tr>"
    HTMLbody = HTMLbody & "</tr>"

    ' 13) Acao necessaria
    HTMLbody = HTMLbody & "<tr>"
    HTMLbody = HTMLbody & "<td><b>Acao necessaria</b></td>"
    HTMLbody = HTMLbody & "<td>" & foundRow.Cells(1, colMap("Detalhamento")).Value & "</td>"
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
        .Subject = "Aprovacao de Custos - Suplementacao de Custos - " & foundRow.Cells(1, 1).Value & " - " & foundRow.Cells(1, 2).Value
        .HTMLbody = HTMLbody & strSignature
        .Display   'Use .Display to just open the email draft
        ' .Send       'Use .Send to send immediately
    End With

    '--- Cleanup
    Set OutMail = Nothing
    Set OutApp = Nothing

    foundRow.Cells(1, colMap("Data")).Value = Date

    MsgBox "Email """ & "Aprovacao de Custos - Suplementacao de Custos - " & foundRow.Cells(1, colMap("Obra")).Value & " - " & foundRow.Cells(1, colMap("Cliente")).Value & """ enviado com sucesso!", vbInformation

CleanExit:
    OptimizeCodeExecution False

End Sub

Sub ClearForm(Optional ShowOnMacroList As Boolean = False)

    OptimizeCodeExecution True

    Dim wsForm As Worksheet
    Dim formMap As Object

    ' Set worksheet reference
    Set wsForm = ThisWorkbook.Sheets("Formulário")
    Set formMap = GetFormFieldsMapping()

    If wsForm.OLEObjects("ComboBoxID").Object.Value = "" Then
        If MsgBox("Esses dados nao foram salvos. Deseja limpa-los mesmo assima", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    ' Populate worksheet with retrieved data
    With wsForm
        .OLEObjects("ComboBoxID").Object.Value = ""
        .OLEObjects("ComboBoxName").Object.Value = ""
        .OLEObjects("ComboBoxName").Width = 123

        ' Read column B values
        .Range(formMap("Obra")).Value = ""
        .Range(formMap("Cliente")).Value = ""
        .Range(formMap("Tipo")).Value = ""
        .Range(formMap("PM")).Value = ""
        .Range(formMap("PEP")).Value = ""
        '.Range(formMap("DR")).Value = ""
        .Range(formMap("Suplementacao")).Value = ""
        .Range(formMap("COT")).Value = ""
        .Range(formMap("Custo Antes")).Value = ""
        .Range(formMap("Custo Depois")).Value = ""
        .Range(formMap("Resultado Inicial")).Value = ""
        .Range(formMap("Resultado Antes")).Value = ""
        .Range(formMap("Resultado Depois")).Value = ""
        .Range(formMap("Provisao")).Value = ""

        ' Read column D values
        .Range(formMap("Descricao")).Value = ""
        .Range(formMap("Justificativa")).Value = ""
        .Range(formMap("Estagio")).Value = ""
        .Range(formMap("Fator")).Value = ""
        .Range(formMap("Detalhamento")).Value = ""
        .Range(formMap("Repasse")).Value = ""
        .Range(formMap("Justificativa Repasse")).Value = ""
        .Range(formMap("Prestador")).Value = ""
        .Range(formMap("Riscos")).Value = ""

        ' Read column F values
        .Range(formMap("Status")).Value = ""
        .Range(formMap("RFP")).Value = ""
        .Range(formMap("Suprimentos")).Value = ""
        .Range(formMap("Pedido")).Value = ""
        .Range(formMap("observacoes")).Value = ""
    End With

    OptimizeCodeExecution False

End Sub

Public Function GetFormFieldsMapping() As Object
    Dim fields As Object
    Set fields = CreateObject("Scripting.Dictionary")

    fields.Add "Obra", "B6"
    fields.Add "Cliente", "B10"
    fields.Add "Tipo", "B14"
    fields.Add "PM", "B18"
    fields.Add "PEP", "B22"
    'fields.Add "DR", "B28"
    fields.Add "Suplementacao", "B28"
    fields.Add "COT", "B32"
    fields.Add "Resultado Inicial", "B36"
    fields.Add "Custo Antes", "B40"
    fields.Add "Resultado Antes", "B44"
    fields.Add "Custo Depois", "B48"
    fields.Add "Resultado Depois", "B52"
    fields.Add "Provisao", "B56"
    
    fields.Add "Descricao", "D6"
    fields.Add "Justificativa", "D10"
    fields.Add "Estagio", "D14"
    fields.Add "Fator", "D18"
    fields.Add "Detalhamento", "D22"
    fields.Add "Repasse", "D26"
    fields.Add "Justificativa Repasse", "D30"
    fields.Add "Prestador", "D34"
    fields.Add "Riscos", "D38"
    
    fields.Add "Status", "F6"
    fields.Add "RFP", "F10"
    fields.Add "Suprimentos", "F14"
    fields.Add "Pedido", "F18"
    fields.Add "observacoes", "F22"

    Set GetFormFieldsMapping = fields
End Function

Public Function GetColumnHeadersMapping() As Object
    Dim headers As Object
    Set headers = CreateObject("Scripting.Dictionary")

    ' Add each header from the provided table to the dictionary,
    ' mapping it to its column position.
    headers.Add "ID", 1
    headers.Add "Obra", headers("ID") + 1
    headers.Add "Cliente", headers("Obra") + 1
    headers.Add "Tipo", headers("Cliente") + 1
    headers.Add "PM", headers("Tipo") + 1
    headers.Add "PEP", headers("PM") + 1
    headers.Add "DR", headers("PEP") + 1
    headers.Add "Suplementacao", headers("DR") + 1
    headers.Add "COT", headers("Suplementacao") + 1
    headers.Add "Resultado Inicial", headers("COT") + 1
    headers.Add "Custo Antes", headers("Resultado Inicial") + 1
    headers.Add "Resultado Antes", headers("Custo Antes") + 1
    headers.Add "Custo Depois", headers("Resultado Antes") + 1
    headers.Add "Resultado Depois", headers("Custo Depois") + 1
    headers.Add "Impacto", headers("Resultado Depois") + 1
    headers.Add "Saldo", headers("Impacto") + 1
    headers.Add "Provisao", headers("Saldo") + 1
    headers.Add "Descricao", headers("Provisao") + 1
    headers.Add "Justificativa", headers("Descricao") + 1
    headers.Add "Estagio", headers("Justificativa") + 1
    headers.Add "Fator", headers("Estagio") + 1
    headers.Add "Detalhamento", headers("Fator") + 1
    headers.Add "Repasse", headers("Detalhamento") + 1
    headers.Add "Justificativa Repasse", headers("Repasse") + 1
    headers.Add "Prestador", headers("Justificativa Repasse") + 1
    headers.Add "Riscos", headers("Prestador") + 1
    headers.Add "Status", headers("Riscos") + 1
    headers.Add "RFP", headers("Status") + 1
    headers.Add "Suprimentos", headers("RFP") + 1
    headers.Add "Pedido", headers("Suprimentos") + 1
    headers.Add "Data", headers("Pedido") + 1
    headers.Add "observacoes", headers("Data") + 1

    Set GetColumnHeadersMapping = headers
End Function

Private Function IsBlankValue(v As Variant) As Boolean
    ' Treat empty or whitespace-only strings as blank; 0 is NOT blank
    If IsError(v) Then
        IsBlankValue = False
    ElseIf IsEmpty(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Trim$(v) = "")
    Else
        IsBlankValue = False
    End If
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

