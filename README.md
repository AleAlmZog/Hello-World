# Hello-World
My first GitHub test repository


Option Compare Database
Option Explicit


Function GeraOFXCitibank(strPath, strQuery, strBankID, strAccID, strOFXName)

On Error GoTo Error_Handler

Dim dbs As Database
Dim rstExport, rstDadosOriginais, rsttmpDados, rstSaldoFinal As Recordset
Dim strSQL, oldFileName, newFileName, strFonte, strDestino, strDataInicio, strDataFim, strDataServ, strSaldoFinal, strDataTransac As String
Dim intCont As Integer


Set dbs = CurrentDb

'O xls gerado pelo banco é na verdade um csv.Copia e muda o sufixo para um que o access abra o arquivo corretamente
strFonte = strPath & strQuery & ".xls"
strDestino = strPath & strQuery & ".csv"
FileCopy strFonte, strDestino

'Limpa os registros de tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
dbs.Execute "DELETE * FROM tmpDadosFormatados"


'Abre tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
strSQL = "SELECT * FROM tmpDadosFormatados"
Set rsttmpDados = dbs.OpenRecordset(strSQL)

'Verifica se ha transações válidas no extrato em processamento:
'seleciona apenas as transações cujo Memo não seja RECURSOS EM C/C', 'SALDO FINAL', ou 'SALDO ANTERIOR'
' as ordena da mais antiga para a mais recente
strSQL = "SELECT Campo1 AS DataTransac, Campo2 AS [Memo], Campo3 AS Valor " & _
        "FROM " & strQuery & _
        " WHERE ([Campo2] Not Like 'RECURSOS EM C/C' And [Campo2] Not Like 'SALDO FINAL' And [Campo2] Not Like 'SALDO ANTERIOR')"

Set rstDadosOriginais = dbs.OpenRecordset(strSQL)

rstDadosOriginais.MoveLast
If rstDadosOriginais.RecordCount = 0 Then
    MsgBox "Nenhuma transação nesta conta no período selecionado", vbInformation + vbOKOnly
    GoTo Exit_Function
End If

'Determina o saldo final: seleciona apenas os registros com "SALDO FINAL" no campo Memo,
'ordena do mais velho ao mais recente, e paga o saldo do mais recente no campo Valor
strSQL = "SELECT Campo1 AS DataTransac, Campo2 AS [Memo], Campo3 AS Valor " & _
        "FROM " & strQuery & _
        " WHERE ([Campo2] Like 'SALDO FINAL')" & _
        " ORDER BY Campo1"

Set rstSaldoFinal = dbs.OpenRecordset(strSQL)
rstSaldoFinal.MoveLast
strSaldoFinal = Replace(Format(rstSaldoFinal.Fields!Valor, "0.00;-0.00"), ",", ".")
rstSaldoFinal.Close
Set rstSaldoFinal = Nothing


'Determina as datas do servidor, inicial e final do extrato
rstDadosOriginais.MoveFirst
strDataInicio = Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd")
rstDadosOriginais.MoveLast
strDataFim = Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd")
strDataServ = strDataFim & "080000"

'o contador intCont serve para deferenciar as transações ocorridas dentro de uma mesma data
'atribuindo letras sucessivas a cada uma a partir de "a" = chr(97), para que o código FITID
'seja composto pela data da transação & sua letra
intCont = 97

'Formata os dados das transações
rsttmpDados.AddNew
rstDadosOriginais.MoveFirst
Do While Not rstDadosOriginais.EOF
    'a variável str DataTransac registra a data da transação que será processada. Ao final do loop
    'vou compará-la com a data da transação seguinte. Se as datas forem as mesmas, intCont=intCont+1
    'e o FITID da transação seguinte será DATA & intCont. Senão, a data da próxima transação
    'é diferente da data da transação anterior, e o intCont volta para 1
    
    strDataTransac = rstDadosOriginais.Fields!DataTransac
   
    With rsttmpDados
        .AddNew
        .Fields!DataTransac = "<DTPOSTED>" & Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd") & "080000"
        If rstDadosOriginais.Fields!Valor < 0 Then
            .Fields!TipoTransac = "<TRNTYPE>DEBIT"
        Else
            .Fields!TipoTransac = "<TRNTYPE>CREDIT"
        End If
        .Fields!Memo = "<MEMO>" & rstDadosOriginais.Fields!Memo
        .Fields!Valor = "<TRNAMT>" & Replace(Format(rstDadosOriginais.Fields!Valor, "0.00;-0.00"), ",", ".")
        .Fields!Fitid = "<FITID>" & Replace(rstDadosOriginais.Fields!DataTransac, "/", "") & "a" & Chr(intCont)
        .Update
        rstDadosOriginais.MoveNext
        
    End With
    
    'comparo a data da nova transação com a data da anterior. Se forem as mesmas, intCont = intCont+1
    'para que o FITID seja DATA & intCont
    If rstDadosOriginais.Fields!DataTransac = strDataTransac Then
        intCont = intCont + 1
    Else
        intCont = 97
    End If
Loop

'este desvio existe por conta de tratamento de erro:
'como no loop acima compara data da transação atual com data da próxima transação
'e este loop roda até EOF, ao chegar EOF nã existe mais data da próxima transação e resulta em erro nº 3021
'sempre que ocorre este número de erro, o tratamento de erro retorna a execução
'para este desvio

Prossegue:
'Insere os dados formatados no campo MEMO que será exportado como OFX

'Limpa tabela com o campo MEMO que será exportado como o ofx ao final da rotina
dbs.Execute "DELETE * FROM TxtAExportar"

'Abre tabela com o campo MEMO que será exportado como o ofx ao final da rotina
strSQL = "SELECT * FROM TxtAExportar"
Set rstExport = dbs.OpenRecordset(strSQL)

'insere o cabeçalho OFX. Ha que se fazer em 2 etapas pois o Access aceita no máximo 25 linhas com CrLf
'Etapa 1:

With rstExport
    .AddNew
    .Fields!TxtAExportar = "OFXHEADER:100" & vbCrLf & _
                            "DATA:OFXSGML" & vbCrLf & _
                            "VERSION:102" & vbCrLf & _
                            "SECURITY:NONE" & vbCrLf & _
                            "ENCODING:USASCII" & vbCrLf & _
                            "CHARSET:1252" & vbCrLf & _
                            "COMPRESSION:NONE" & vbCrLf & _
                            "OLDFILEUID:NONE" & vbCrLf & _
                            "NEWFILEUID:NONE" & vbCrLf & _
                            "<OFX>" & vbCrLf & _
                            "<SIGNONMSGSRSV1>" & vbCrLf & _
                            "<SONRS>" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<DTSERVER>" & strDataServ & vbCrLf & _
                            "<LANGUAGE>POR" & vbCrLf & _
                            "</SONRS>"
    .Update
    'Etapa 2
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
                            "</SIGNONMSGSRSV1>" & vbCrLf & _
                            "<BANKMSGSRSV1>" & vbCrLf & _
                            "<STMTTRNRS>" & vbCrLf & _
                            "<TRNUID>1001" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<STMTRS>" & vbCrLf & _
                            "<CURDEF>BRL" & vbCrLf & _
                            "<BANKACCTFROM>" & vbCrLf & _
                            "<BANKID>" & strBankID & vbCrLf & _
                            "<ACCTID>" & strAccID & vbCrLf & _
                            "<ACCTTYPE>CHECKING" & vbCrLf & _
                            "</BANKACCTFROM>" & vbCrLf & _
                            "<BANKTRANLIST>" & vbCrLf & _
                            "<DTSTART>" & strDataInicio & vbCrLf & _
                            "<DTEND>" & strDataFim & vbCrLf
    .Update
End With


        
'insere os dados das transações
With rsttmpDados
    .MoveFirst
    Do While Not .EOF
        rstExport.MoveFirst
        rstExport.Edit
        rstExport.Fields!TxtAExportar = rstExport.Fields!TxtAExportar & vbCrLf & _
            "<STMTTRN>" & vbCrLf & _
            .Fields!TipoTransac & vbCrLf & _
            .Fields!DataTransac & vbCrLf & _
            .Fields!Valor & vbCrLf & _
            .Fields!Fitid & vbCrLf & _
            .Fields!Memo & vbCrLf & _
            "</STMTTRN>" & vbCrLf
        rstExport.Update
        .MoveNext
    Loop
End With
    
With rstExport
    'insere o rodapé WPL
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
        "</BANKTRANLIST>" & vbCrLf & _
        "<LEDGERBAL>" & vbCrLf & _
        "<BALAMT>" & strSaldoFinal & vbCrLf & _
        "<DTASOF>" & strDataFim & vbCrLf & _
        "</LEDGERBAL>" & vbCrLf & _
        "</STMTRS>" & vbCrLf & _
        "</STMTTRNRS>" & vbCrLf & _
        "</BANKMSGSRSV1>" & vbCrLf & _
        "</OFX>"
    .Update
End With

rstExport.Close
rstDadosOriginais.Close
rsttmpDados.Close

Set rstExport = Nothing
Set rstDadosOriginais = Nothing
Set rsttmpDados = Nothing
    
Set dbs = Nothing
    
    'crio o nome do .txt a exportar
    oldFileName = strPath & strOFXName & ".txt"
            
    'crio com sufixo .ofx o nome do arquivo para renomear o arquivo exportado com sufixo .txt
    newFileName = strPath & strOFXName & ".ofx"
            
    'exporta .txt
    DoCmd.TransferText acExportDelim, "specTxtAExportar", "TxtAExportar", oldFileName
    
    'Checa se o arquivo com o mesmo nome jé existe
    If FileOrDirExists(newFileName) = True Then
        'Se já existir, deleta
        Kill newFileName
    End If
    
    'Renomeio .txt para .ofx
    Name oldFileName As newFileName
    
'Apaga .csv
Kill strDestino

'avisa que rotina foiexecutada a contento
Debug.Print MsgBox(strOFXName & " exportado com sucesso", vbOKOnly)

Exit_Function:
    Exit Function

Error_Handler:
    'erro 3021: no loop cujo fim é EOF eu comparo a data da transação do registro anterior com a data
    'da transação onde está o ponteiro. Quando chega em EOF, não ha mais data no ponteiro para comparar com a anterior
    'Assim, este erro significa apenas que o loop acabou e devolvo a rotina para o ponto onde deve continuar
    If Err.Number = 3021 Then
        GoTo Prossegue
    Else
        MsgBox Err.Description & Err.Number
        Resume Exit_Function
    End If

End Function

Function GeraOFXSantanderContaCorrente(strPath, strQuery, strBankID, strAccID, strOFXName)

On Error GoTo Error_Handler

Dim dbs As Database
Dim rstExport, rstDadosOriginais, rsttmpDados, rstSaldoFinal As Recordset
Dim strSQL, oldFileName, newFileName, strFonte, strDestino, strDataInicio, strDataFim, strDataServ, strSaldoFinal, strDataTransac As String
Dim intCont As Integer


Set dbs = CurrentDb

'O xls gerado pelo banco é na verdade um csv.Copia e muda o sufixo para um que o access abra o arquivo corretamente
strFonte = strPath & strQuery & ".xls"

'Limpa os registros de tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
dbs.Execute "DELETE * FROM tmpDadosFormatados"


'Abre tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
strSQL = "SELECT * FROM tmpDadosFormatados"
Set rsttmpDados = dbs.OpenRecordset(strSQL)

'Verifica se ha transações válidas no extrato em processamento:
strSQL = "SELECT SantanderContaCorrenteLimpo.* FROM SantanderContaCorrenteLimpo"

Set rstDadosOriginais = dbs.OpenRecordset(strSQL)

rstDadosOriginais.MoveLast
If rstDadosOriginais.RecordCount = 0 Then
    MsgBox "Nenhuma transação nesta conta no período selecionado", vbInformation + vbOKOnly
    GoTo Exit_Function
End If

'Determina o saldo final: seleciona o último registro do campo Saldo,
strSaldoFinal = Replace(Format(rstDadosOriginais.Fields!Saldo, "0.00;-0.00"), ",", ".")

'Determina as datas do servidor, inicial e final do extrato
rstDadosOriginais.MoveFirst
strDataInicio = Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd")
rstDadosOriginais.MoveLast
strDataFim = Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd")
strDataServ = strDataFim & "080000"

'o contador intCont serve para deferenciar as transações ocorridas dentro de uma mesma data
'atribuindo letras sucessivas a cada uma a partir de "a" = chr(97), para que o código FITID
'seja composto pela data da transação & sua letra
intCont = 97

'Formata os dados das transações
rsttmpDados.AddNew
rstDadosOriginais.MoveFirst
Do While Not rstDadosOriginais.EOF
    'a variável str DataTransac registra a data da transação que será processada. Ao final do loop
    'vou compará-la com a data da transação seguinte. Se as datas forem as mesmas, intCont=intCont+1
    'e o FITID da transação seguinte será DATA & intCont. Senão, a data da próxima transação
    'é diferente da data da transação anterior, e o intCont volta para 1
    
    strDataTransac = rstDadosOriginais.Fields!DataTransac
   
    With rsttmpDados
        .AddNew
        .Fields!DataTransac = "<DTPOSTED>" & Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd") & "080000"
        If rstDadosOriginais.Fields!Valor < 0 Then
            .Fields!TipoTransac = "<TRNTYPE>DEBIT"
        Else
            .Fields!TipoTransac = "<TRNTYPE>CREDIT"
        End If
        .Fields!Memo = "<MEMO>" & Replace(rstDadosOriginais.Fields!Memo, "&", "e")
        .Fields!Valor = "<TRNAMT>" & Replace(Format(rstDadosOriginais.Fields!Valor, "0.00;-0.00"), ",", ".")
        .Fields!Fitid = "<FITID>" & Replace(rstDadosOriginais.Fields!DataTransac, "/", "") & "a" & Chr(intCont)
        .Update
        rstDadosOriginais.MoveNext
        
    End With
    
    'comparo a data da nova transação com a data da anterior. Se forem as mesmas, intCont = intCont+1
    'para que o FITID seja DATA & intCont
    If rstDadosOriginais.Fields!DataTransac = strDataTransac Then
        intCont = intCont + 1
    Else
        intCont = 97
    End If
Loop

'este desvio existe por conta de tratamento de erro:
'como no loop acima compara data da transação atual com data da próxima transação
'e este loop roda até EOF, ao chegar EOF nã existe mais data da próxima transação e resulta em erro nº 3021
'sempre que ocorre este número de erro, o tratamento de erro retorna a execução
'para este desvio

Prossegue:
'Insere os dados formatados no campo MEMO que será exportado como OFX

'Limpa tabela com o campo MEMO que será exportado como o ofx ao final da rotina
dbs.Execute "DELETE * FROM TxtAExportar"

'Abre tabela com o campo MEMO que será exportado como o ofx ao final da rotina
strSQL = "SELECT * FROM TxtAExportar"
Set rstExport = dbs.OpenRecordset(strSQL)

'insere o cabeçalho OFX. Ha que se fazer em 2 etapas pois o Access aceita no máximo 25 linhas com CrLf
'Etapa 1:

With rstExport
    .AddNew
    .Fields!TxtAExportar = "OFXHEADER:100" & vbCrLf & _
                            "DATA:OFXSGML" & vbCrLf & _
                            "VERSION:102" & vbCrLf & _
                            "SECURITY:NONE" & vbCrLf & _
                            "ENCODING:USASCII" & vbCrLf & _
                            "CHARSET:1252" & vbCrLf & _
                            "COMPRESSION:NONE" & vbCrLf & _
                            "OLDFILEUID:NONE" & vbCrLf & _
                            "NEWFILEUID:NONE" & vbCrLf & _
                            "<OFX>" & vbCrLf & _
                            "<SIGNONMSGSRSV1>" & vbCrLf & _
                            "<SONRS>" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<DTSERVER>" & strDataServ & vbCrLf & _
                            "<LANGUAGE>POR" & vbCrLf & _
                            "</SONRS>"
    .Update
    'Etapa 2
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
                            "</SIGNONMSGSRSV1>" & vbCrLf & _
                            "<BANKMSGSRSV1>" & vbCrLf & _
                            "<STMTTRNRS>" & vbCrLf & _
                            "<TRNUID>1001" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<STMTRS>" & vbCrLf & _
                            "<CURDEF>EUR" & vbCrLf & _
                            "<BANKACCTFROM>" & vbCrLf & _
                            "<BANKID>" & strBankID & vbCrLf & _
                            "<ACCTID>" & strAccID & vbCrLf & _
                            "<ACCTTYPE>CHECKING" & vbCrLf & _
                            "</BANKACCTFROM>" & vbCrLf & _
                            "<BANKTRANLIST>" & vbCrLf & _
                            "<DTSTART>" & strDataInicio & vbCrLf & _
                            "<DTEND>" & strDataFim & vbCrLf
    .Update
End With


        
'insere os dados das transações
With rsttmpDados
    .MoveFirst
    Do While Not .EOF
        rstExport.MoveFirst
        rstExport.Edit
        rstExport.Fields!TxtAExportar = rstExport.Fields!TxtAExportar & vbCrLf & _
            "<STMTTRN>" & vbCrLf & _
            .Fields!TipoTransac & vbCrLf & _
            .Fields!DataTransac & vbCrLf & _
            .Fields!Valor & vbCrLf & _
            .Fields!Fitid & vbCrLf & _
            .Fields!Memo & vbCrLf & _
            "</STMTTRN>" & vbCrLf
        rstExport.Update
        .MoveNext
    Loop
End With
    
With rstExport
    'insere o rodapé WPL
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
        "</BANKTRANLIST>" & vbCrLf & _
        "<LEDGERBAL>" & vbCrLf & _
        "<BALAMT>" & strSaldoFinal & vbCrLf & _
        "<DTASOF>" & strDataFim & vbCrLf & _
        "</LEDGERBAL>" & vbCrLf & _
        "</STMTRS>" & vbCrLf & _
        "</STMTTRNRS>" & vbCrLf & _
        "</BANKMSGSRSV1>" & vbCrLf & _
        "</OFX>"
    .Update
End With

rstExport.Close
rstDadosOriginais.Close
rsttmpDados.Close

Set rstExport = Nothing
Set rstDadosOriginais = Nothing
Set rsttmpDados = Nothing
    
Set dbs = Nothing
    
    'crio o nome do .txt a exportar
    oldFileName = strPath & strOFXName & ".txt"
            
    'crio com sufixo .ofx o nome do arquivo para renomear o arquivo exportado com sufixo .txt
    newFileName = strPath & strOFXName & ".ofx"
            
    'exporta .txt
    DoCmd.TransferText acExportDelim, "specTxtAExportar", "TxtAExportar", oldFileName
    
    'Checa se o arquivo com o mesmo nome jé existe
    If FileOrDirExists(newFileName) = True Then
        'Se já existir, deleta
        Kill newFileName
    End If
    
    'Renomeio .txt para .ofx
    Name oldFileName As newFileName
    
'avisa que rotina foiexecutada a contento
Debug.Print MsgBox(strOFXName & " exportado com sucesso", vbOKOnly)

Exit_Function:
    Exit Function

Error_Handler:
    'erro 3021: no loop cujo fim é EOF eu comparo a data da transação do registro anterior com a data
    'da transação onde está o ponteiro. Quando chega em EOF, não ha mais data no ponteiro para comparar com a anterior
    'Assim, este erro significa apenas que o loop acabou e devolvo a rotina para o ponto onde deve continuar
    If Err.Number = 3021 Then
        GoTo Prossegue
    Else
        MsgBox Err.Description & Err.Number
        Resume Exit_Function
    End If

End Function

Function GeraOFXSantanderMastercard(strPath, strQuery, strBankID, strAccID, strOFXName)

On Error GoTo Error_Handler

Dim dbs As Database
Dim rstExport, rstDadosOriginais, rsttmpDados, rstSaldoFinal As Recordset
Dim strSQL, oldFileName, newFileName, strFonte, strDestino, strDataInicio, strDataFim, strDataServ, strSaldoFinal, strDataTransac As String
Dim intCont As Integer


Set dbs = CurrentDb

'O xls gerado pelo banco é na verdade um csv.Copia e muda o sufixo para um que o access abra o arquivo corretamente
strFonte = strPath & strQuery & ".xls"

'Limpa os registros de tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
dbs.Execute "DELETE * FROM tmpDadosFormatados"


'Abre tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
strSQL = "SELECT * FROM tmpDadosFormatados"
Set rsttmpDados = dbs.OpenRecordset(strSQL)

'Verifica se ha transações válidas no extrato em processamento:
strSQL = "SELECT SantanderMastercardLimpo.* FROM SantanderMastercardLimpo"

Set rstDadosOriginais = dbs.OpenRecordset(strSQL)

rstDadosOriginais.MoveLast
If rstDadosOriginais.RecordCount = 0 Then
    MsgBox "Nenhuma transação nesta conta no período selecionado", vbInformation + vbOKOnly
    GoTo Exit_Function
End If

'Determina o saldo final: o banco não exporta esta informação para cartões de crédito, assumo que será sempre 0
strSaldoFinal = "0.00"

'Determina as datas do servidor, inicial e final do extrato
rstDadosOriginais.MoveFirst
strDataInicio = Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd")
rstDadosOriginais.MoveLast
strDataFim = Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd")
strDataServ = strDataFim & "080000"

'o contador intCont serve para deferenciar as transações ocorridas dentro de uma mesma data
'atribuindo letras sucessivas a cada uma a partir de "a" = chr(97), para que o código FITID
'seja composto pela data da transação & sua letra
intCont = 97

'Formata os dados das transações
rsttmpDados.AddNew
rstDadosOriginais.MoveFirst
Do While Not rstDadosOriginais.EOF
    'a variável str DataTransac registra a data da transação que será processada. Ao final do loop
    'vou compará-la com a data da transação seguinte. Se as datas forem as mesmas, intCont=intCont+1
    'e o FITID da transação seguinte será DATA & intCont. Senão, a data da próxima transação
    'é diferente da data da transação anterior, e o intCont volta para 1
    
    strDataTransac = rstDadosOriginais.Fields!DataTransac
   
    With rsttmpDados
        .AddNew
        .Fields!DataTransac = "<DTPOSTED>" & Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd") & "080000"
        If rstDadosOriginais.Fields!DebCred = "D" Then
            .Fields!TipoTransac = "<TRNTYPE>DEBIT"
            .Fields!Valor = "<TRNAMT>" & "-" & Replace(Format(rstDadosOriginais.Fields!Valor, "0.00;-0.00"), ",", ".")
        Else
            .Fields!TipoTransac = "<TRNTYPE>CREDIT"
            .Fields!Valor = "<TRNAMT>" & Replace(Format(rstDadosOriginais.Fields!Valor, "0.00;-0.00"), ",", ".")
        End If
        .Fields!Memo = "<MEMO>" & Replace(rstDadosOriginais.Fields!Memo, "&", "e")
        .Fields!Fitid = "<FITID>" & Replace(rstDadosOriginais.Fields!DataTransac, "/", "") & "a" & Chr(intCont)
        .Update
        rstDadosOriginais.MoveNext
        
    End With
    
    'comparo a data da nova transação com a data da anterior. Se forem as mesmas, intCont = intCont+1
    'para que o FITID seja DATA & intCont
    If rstDadosOriginais.Fields!DataTransac = strDataTransac Then
        intCont = intCont + 1
    Else
        intCont = 97
    End If
Loop

'este desvio existe por conta de tratamento de erro:
'como no loop acima compara data da transação atual com data da próxima transação
'e este loop roda até EOF, ao chegar EOF nã existe mais data da próxima transação e resulta em erro nº 3021
'sempre que ocorre este número de erro, o tratamento de erro retorna a execução
'para este desvio

Prossegue:
'Insere os dados formatados no campo MEMO que será exportado como OFX

'Limpa tabela com o campo MEMO que será exportado como o ofx ao final da rotina
dbs.Execute "DELETE * FROM TxtAExportar"

'Abre tabela com o campo MEMO que será exportado como o ofx ao final da rotina
strSQL = "SELECT * FROM TxtAExportar"
Set rstExport = dbs.OpenRecordset(strSQL)

'insere o cabeçalho OFX. Ha que se fazer em 2 etapas pois o Access aceita no máximo 25 linhas com CrLf
'Etapa 1:

With rstExport
    .AddNew
    .Fields!TxtAExportar = "OFXHEADER:100" & vbCrLf & _
                            "DATA:OFXSGML" & vbCrLf & _
                            "VERSION:102" & vbCrLf & _
                            "SECURITY:NONE" & vbCrLf & _
                            "ENCODING:USASCII" & vbCrLf & _
                            "CHARSET:1252" & vbCrLf & _
                            "COMPRESSION:NONE" & vbCrLf & _
                            "OLDFILEUID:NONE" & vbCrLf & _
                            "NEWFILEUID:NONE" & vbCrLf & _
                            "<OFX>" & vbCrLf & _
                            "<SIGNONMSGSRSV1>" & vbCrLf & _
                            "<SONRS>" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<DTSERVER>" & strDataServ & vbCrLf & _
                            "<LANGUAGE>POR" & vbCrLf & _
                            "</SONRS>"
    .Update
    'Etapa 2
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
                            "</SIGNONMSGSRSV1>" & vbCrLf & _
                            "<BANKMSGSRSV1>" & vbCrLf & _
                            "<STMTTRNRS>" & vbCrLf & _
                            "<TRNUID>1001" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<STMTRS>" & vbCrLf & _
                            "<CURDEF>EUR" & vbCrLf & _
                            "<BANKACCTFROM>" & vbCrLf & _
                            "<BANKID>" & strBankID & vbCrLf & _
                            "<ACCTID>" & strAccID & vbCrLf & _
                            "<ACCTTYPE>CHECKING" & vbCrLf & _
                            "</BANKACCTFROM>" & vbCrLf & _
                            "<BANKTRANLIST>" & vbCrLf & _
                            "<DTSTART>" & strDataInicio & vbCrLf & _
                            "<DTEND>" & strDataFim & vbCrLf
    .Update
End With


        
'insere os dados das transações
With rsttmpDados
    .MoveFirst
    Do While Not .EOF
        rstExport.MoveFirst
        rstExport.Edit
        rstExport.Fields!TxtAExportar = rstExport.Fields!TxtAExportar & vbCrLf & _
            "<STMTTRN>" & vbCrLf & _
            .Fields!TipoTransac & vbCrLf & _
            .Fields!DataTransac & vbCrLf & _
            .Fields!Valor & vbCrLf & _
            .Fields!Fitid & vbCrLf & _
            .Fields!Memo & vbCrLf & _
            "</STMTTRN>" & vbCrLf
        rstExport.Update
        .MoveNext
    Loop
End With
    
With rstExport
    'insere o rodapé WPL
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
        "</BANKTRANLIST>" & vbCrLf & _
        "<LEDGERBAL>" & vbCrLf & _
        "<BALAMT>" & strSaldoFinal & vbCrLf & _
        "<DTASOF>" & strDataFim & vbCrLf & _
        "</LEDGERBAL>" & vbCrLf & _
        "</STMTRS>" & vbCrLf & _
        "</STMTTRNRS>" & vbCrLf & _
        "</BANKMSGSRSV1>" & vbCrLf & _
        "</OFX>"
    .Update
End With

rstExport.Close
rstDadosOriginais.Close
rsttmpDados.Close

Set rstExport = Nothing
Set rstDadosOriginais = Nothing
Set rsttmpDados = Nothing
    
Set dbs = Nothing
    
    'crio o nome do .txt a exportar
    oldFileName = strPath & strOFXName & ".txt"
            
    'crio com sufixo .ofx o nome do arquivo para renomear o arquivo exportado com sufixo .txt
    newFileName = strPath & strOFXName & ".ofx"
            
    'exporta .txt
    DoCmd.TransferText acExportDelim, "specTxtAExportar", "TxtAExportar", oldFileName
    
    'Checa se o arquivo com o mesmo nome jé existe
    If FileOrDirExists(newFileName) = True Then
        'Se já existir, deleta
        Kill newFileName
    End If
    
    'Renomeio .txt para .ofx
    Name oldFileName As newFileName
    
'avisa que rotina foiexecutada a contento
Debug.Print MsgBox(strOFXName & " exportado com sucesso", vbOKOnly)

Exit_Function:
    Exit Function

Error_Handler:
    'erro 3021: no loop cujo fim é EOF eu comparo a data da transação do registro anterior com a data
    'da transação onde está o ponteiro. Quando chega em EOF, não ha mais data no ponteiro para comparar com a anterior
    'Assim, este erro significa apenas que o loop acabou e devolvo a rotina para o ponto onde deve continuar
    If Err.Number = 3021 Then
        GoTo Prossegue
    Else
        MsgBox Err.Description & Err.Number
        Resume Exit_Function
    End If

End Function


Function GeraOFXCSHG(strPath, strQuery, strBankID, strAccID, strOFXName)

On Error GoTo Error_Handler

Dim dbs As Database
Dim rstExport, rstDadosOriginais, rsttmpDados, rstSaldoFinal As Recordset
Dim strSQL, oldFileName, newFileName, strFonte, strDestino, strDataInicio, strDataFim, strDataServ, strSaldoFinal, strDataTransac As String
Dim intCont As Integer


Set dbs = CurrentDb

'O xls gerado pelo banco é na verdade um csv.Copia e muda o sufixo para um que o access abra o arquivo corretamente
strFonte = strPath & strQuery & ".xls"

'Limpa os registros de tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
dbs.Execute "DELETE * FROM tmpDadosFormatados"


'Abre tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
strSQL = "SELECT * FROM tmpDadosFormatados"
Set rsttmpDados = dbs.OpenRecordset(strSQL)

'Verifica se ha transações válidas no extrato em processamento:
strSQL = "SELECT CSHGAndyLimpo.* FROM CSHGAndyLimpo"

Set rstDadosOriginais = dbs.OpenRecordset(strSQL)

rstDadosOriginais.MoveLast
If rstDadosOriginais.RecordCount = 0 Then
    MsgBox "Nenhuma transação nesta conta no período selecionado", vbInformation + vbOKOnly
    GoTo Exit_Function
End If

'Determina o saldo final: seleciona o último registro do campo Saldo,
'strSaldoFinal = Replace(Format(rstDadosOriginais.Fields!Saldo, "0.00;-0.00"), ",", ".")

'Determina as datas do servidor, inicial e final do extrato
rstDadosOriginais.MoveFirst
strDataInicio = Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd")
rstDadosOriginais.MoveLast
strDataFim = Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd")
strDataServ = strDataFim & "080000"

'o contador intCont serve para deferenciar as transações ocorridas dentro de uma mesma data
'atribuindo letras sucessivas a cada uma a partir de "a" = chr(97), para que o código FITID
'seja composto pela data da transação & sua letra
intCont = 97

'Formata os dados das transações
rsttmpDados.AddNew
rstDadosOriginais.MoveFirst
Do While Not rstDadosOriginais.EOF
    'a variável str DataTransac registra a data da transação que será processada. Ao final do loop
    'vou compará-la com a data da transação seguinte. Se as datas forem as mesmas, intCont=intCont+1
    'e o FITID da transação seguinte será DATA & intCont. Senão, a data da próxima transação
    'é diferente da data da transação anterior, e o intCont volta para 1
    
    strDataTransac = rstDadosOriginais.Fields!DataTransac
   
    With rsttmpDados
        .AddNew
        .Fields!DataTransac = "<DTPOSTED>" & Format((rstDadosOriginais.Fields!DataTransac), "yyyymmdd") & "080000"
        If rstDadosOriginais.Fields!Valor < 0 Then
            .Fields!TipoTransac = "<TRNTYPE>DEBIT"
        Else
            .Fields!TipoTransac = "<TRNTYPE>CREDIT"
        End If
        .Fields!Memo = "<MEMO>" & Replace(rstDadosOriginais.Fields!Memo, "&", "e")
        .Fields!Valor = "<TRNAMT>" & Replace(Format(rstDadosOriginais.Fields!Valor, "0.00;-0.00"), ",", ".")
        .Fields!Fitid = "<FITID>" & Replace(rstDadosOriginais.Fields!DataTransac, "/", "") & "a" & Chr(intCont)
        .Update
        rstDadosOriginais.MoveNext
        
    End With
    
    'comparo a data da nova transação com a data da anterior. Se forem as mesmas, intCont = intCont+1
    'para que o FITID seja DATA & intCont
    If rstDadosOriginais.Fields!DataTransac = strDataTransac Then
        intCont = intCont + 1
    Else
        intCont = 97
    End If
Loop

'este desvio existe por conta de tratamento de erro:
'como no loop acima compara data da transação atual com data da próxima transação
'e este loop roda até EOF, ao chegar EOF nã existe mais data da próxima transação e resulta em erro nº 3021
'sempre que ocorre este número de erro, o tratamento de erro retorna a execução
'para este desvio

Prossegue:
'Insere os dados formatados no campo MEMO que será exportado como OFX

'Limpa tabela com o campo MEMO que será exportado como o ofx ao final da rotina
dbs.Execute "DELETE * FROM TxtAExportar"

'Abre tabela com o campo MEMO que será exportado como o ofx ao final da rotina
strSQL = "SELECT * FROM TxtAExportar"
Set rstExport = dbs.OpenRecordset(strSQL)

'insere o cabeçalho OFX. Ha que se fazer em 2 etapas pois o Access aceita no máximo 25 linhas com CrLf
'Etapa 1:

With rstExport
    .AddNew
    .Fields!TxtAExportar = "OFXHEADER:100" & vbCrLf & _
                            "DATA:OFXSGML" & vbCrLf & _
                            "VERSION:102" & vbCrLf & _
                            "SECURITY:NONE" & vbCrLf & _
                            "ENCODING:USASCII" & vbCrLf & _
                            "CHARSET:1252" & vbCrLf & _
                            "COMPRESSION:NONE" & vbCrLf & _
                            "OLDFILEUID:NONE" & vbCrLf & _
                            "NEWFILEUID:NONE" & vbCrLf & _
                            "<OFX>" & vbCrLf & _
                            "<SIGNONMSGSRSV1>" & vbCrLf & _
                            "<SONRS>" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<DTSERVER>" & strDataServ & vbCrLf & _
                            "<LANGUAGE>POR" & vbCrLf & _
                            "</SONRS>"
    .Update
    'Etapa 2
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
                            "</SIGNONMSGSRSV1>" & vbCrLf & _
                            "<BANKMSGSRSV1>" & vbCrLf & _
                            "<STMTTRNRS>" & vbCrLf & _
                            "<TRNUID>1001" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<STMTRS>" & vbCrLf & _
                            "<CURDEF>BRL" & vbCrLf & _
                            "<BANKACCTFROM>" & vbCrLf & _
                            "<BANKID>" & strBankID & vbCrLf & _
                            "<ACCTID>" & strAccID & vbCrLf & _
                            "<ACCTTYPE>CHECKING" & vbCrLf & _
                            "</BANKACCTFROM>" & vbCrLf & _
                            "<BANKTRANLIST>" & vbCrLf & _
                            "<DTSTART>" & strDataInicio & vbCrLf & _
                            "<DTEND>" & strDataFim & vbCrLf
    .Update
End With


        
'insere os dados das transações
With rsttmpDados
    .MoveFirst
    Do While Not .EOF
        rstExport.MoveFirst
        rstExport.Edit
        rstExport.Fields!TxtAExportar = rstExport.Fields!TxtAExportar & vbCrLf & _
            "<STMTTRN>" & vbCrLf & _
            .Fields!TipoTransac & vbCrLf & _
            .Fields!DataTransac & vbCrLf & _
            .Fields!Valor & vbCrLf & _
            .Fields!Fitid & vbCrLf & _
            .Fields!Memo & vbCrLf & _
            "</STMTTRN>" & vbCrLf
        rstExport.Update
        .MoveNext
    Loop
End With
    
With rstExport
    'insere o rodapé WPL
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
        "</BANKTRANLIST>" & vbCrLf & _
        "<LEDGERBAL>" & vbCrLf & _
        "<BALAMT>" & "0.00" & vbCrLf & _
        "<DTASOF>" & strDataFim & vbCrLf & _
        "</LEDGERBAL>" & vbCrLf & _
        "</STMTRS>" & vbCrLf & _
        "</STMTTRNRS>" & vbCrLf & _
        "</BANKMSGSRSV1>" & vbCrLf & _
        "</OFX>"
    .Update
End With

rstExport.Close
rstDadosOriginais.Close
rsttmpDados.Close

Set rstExport = Nothing
Set rstDadosOriginais = Nothing
Set rsttmpDados = Nothing
    
Set dbs = Nothing
    
    'crio o nome do .txt a exportar
    oldFileName = strPath & strOFXName & ".txt"
            
    'crio com sufixo .ofx o nome do arquivo para renomear o arquivo exportado com sufixo .txt
    newFileName = strPath & strOFXName & ".ofx"
            
    'exporta .txt
    DoCmd.TransferText acExportDelim, "specTxtAExportar", "TxtAExportar", oldFileName
    
    'Checa se o arquivo com o mesmo nome jé existe
    If FileOrDirExists(newFileName) = True Then
        'Se já existir, deleta
        Kill newFileName
    End If
    
    'Renomeio .txt para .ofx
    Name oldFileName As newFileName
    
'avisa que rotina foiexecutada a contento
Debug.Print MsgBox(strOFXName & " exportado com sucesso", vbOKOnly)

Exit_Function:
    Exit Function

Error_Handler:
    'erro 3021: no loop cujo fim é EOF eu comparo a data da transação do registro anterior com a data
    'da transação onde está o ponteiro. Quando chega em EOF, não ha mais data no ponteiro para comparar com a anterior
    'Assim, este erro significa apenas que o loop acabou e devolvo a rotina para o ponto onde deve continuar
    If Err.Number = 3021 Then
        GoTo Prossegue
    Else
        MsgBox Err.Description & Err.Number
        Resume Exit_Function
    End If

End Function

Function GeraOFXLeumi(strPath, strQuery, strBankID, strAccID, strOFXName)

On Error GoTo Error_Handler

Dim dbs As Database
Dim rstExport, rstDadosOriginais, rsttmpDados, rstSaldoFinal As Recordset
Dim strSQL, oldFileName, newFileName, strFonte, strDestino, strDataInicio, strDataFim, strDataServ, strSaldoFinal, strDataTransac As String
Dim intCont As Integer


Set dbs = CurrentDb

'Limpa os registros de tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
dbs.Execute "DELETE * FROM tmpDadosFormatados"


'Abre tmpDadosFormatados, tabela que receberá os dados para o .ofx formatados conforme protocolo Microsoft
strSQL = "SELECT * FROM tmpDadosFormatados"
Set rsttmpDados = dbs.OpenRecordset(strSQL)

'Verifica se ha transações válidas no extrato em processamento:
strSQL = "SELECT LeumiCSV.* FROM LeumiCSV"

Set rstDadosOriginais = dbs.OpenRecordset(strSQL)

rstDadosOriginais.MoveLast
If rstDadosOriginais.RecordCount = 0 Then
    MsgBox "Nenhuma transação nesta conta no período selecionado", vbInformation + vbOKOnly
    GoTo Exit_Function
End If

'Determina o saldo final: seleciona o primeiro registro do campo Saldo,
rstDadosOriginais.MoveFirst
rstDadosOriginais.Move 2
strSaldoFinal = rstDadosOriginais.Fields!Summary

'Determina as datas do servidor, inicial e final do extrato
strDataFim = Format((rstDadosOriginais.Fields![Post Date]), "yyyymmdd")
strDataServ = strDataFim & "080000"
rstDadosOriginais.Move 1
strDataInicio = Format((rstDadosOriginais.Fields![Post Date]), "yyyymmdd")

'o contador intCont serve para deferenciar as transações ocorridas dentro de uma mesma data
'atribuindo letras sucessivas a cada uma a partir de "a" = chr(97), para que o código FITID
'seja composto pela data da transação & sua letra
intCont = 97

'Formata os dados das transações
rsttmpDados.AddNew

Do While Not rstDadosOriginais.EOF
    'a variável str DataTransac registra a data da transação que será processada. Ao final do loop
    'vou compará-la com a data da transação seguinte. Se as datas forem as mesmas, intCont=intCont+1
    'e o FITID da transação seguinte será DATA & intCont. Senão, a data da próxima transação
    'é diferente da data da transação anterior, e o intCont volta para 1
    
    strDataTransac = rstDadosOriginais.Fields![Post Date]
   
    With rsttmpDados
        .AddNew
        .Fields!DataTransac = "<DTPOSTED>" & Format((rstDadosOriginais.Fields![Post Date]), "yyyymmdd") & "080000"
        If IsNull(rstDadosOriginais.Fields!Debit) = False Then
            .Fields!TipoTransac = "<TRNTYPE>DEBIT"
        Else
            .Fields!TipoTransac = "<TRNTYPE>CREDIT"
        End If
        .Fields!Memo = "<MEMO>" & Replace(rstDadosOriginais.Fields!Description, "&", "e") & " " & _
            Replace(rstDadosOriginais.Fields!Text, "&", "e")
        If IsNull(rstDadosOriginais.Fields!Debit) = False Then
            .Fields!Valor = "<TRNAMT>-" & rstDadosOriginais.Fields!Debit
        Else
            .Fields!Valor = "<TRNAMT>" & rstDadosOriginais.Fields!Credit
        End If
        .Fields!Fitid = "<FITID>" & Replace(rstDadosOriginais.Fields![Post Date], "/", "") & "a" & Chr(intCont)
        .Update
        rstDadosOriginais.MoveNext
        
    End With
    
    'comparo a data da nova transação com a data da anterior. Se forem as mesmas, intCont = intCont+1
    'para que o FITID seja DATA & intCont
    If rstDadosOriginais.Fields![Post Date] = strDataTransac Then
        intCont = intCont + 1
    Else
        intCont = 97
    End If
Loop

'este desvio existe por conta de tratamento de erro:
'como no loop acima compara data da transação atual com data da próxima transação
'e este loop roda até EOF, ao chegar EOF nã existe mais data da próxima transação e resulta em erro nº 3021
'sempre que ocorre este número de erro, o tratamento de erro retorna a execução
'para este desvio

Prossegue:
'Insere os dados formatados no campo MEMO que será exportado como OFX

'Limpa tabela com o campo MEMO que será exportado como o ofx ao final da rotina
dbs.Execute "DELETE * FROM TxtAExportar"

'Abre tabela com o campo MEMO que será exportado como o ofx ao final da rotina
strSQL = "SELECT * FROM TxtAExportar"
Set rstExport = dbs.OpenRecordset(strSQL)

'insere o cabeçalho OFX. Ha que se fazer em 2 etapas pois o Access aceita no máximo 25 linhas com CrLf
'Etapa 1:

With rstExport
    .AddNew
    .Fields!TxtAExportar = "OFXHEADER:100" & vbCrLf & _
                            "DATA:OFXSGML" & vbCrLf & _
                            "VERSION:102" & vbCrLf & _
                            "SECURITY:NONE" & vbCrLf & _
                            "ENCODING:USASCII" & vbCrLf & _
                            "CHARSET:1252" & vbCrLf & _
                            "COMPRESSION:NONE" & vbCrLf & _
                            "OLDFILEUID:NONE" & vbCrLf & _
                            "NEWFILEUID:NONE" & vbCrLf & _
                            "<OFX>" & vbCrLf & _
                            "<SIGNONMSGSRSV1>" & vbCrLf & _
                            "<SONRS>" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<DTSERVER>" & strDataServ & vbCrLf & _
                            "<LANGUAGE>POR" & vbCrLf & _
                            "</SONRS>"
    .Update
    'Etapa 2
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
                            "</SIGNONMSGSRSV1>" & vbCrLf & _
                            "<BANKMSGSRSV1>" & vbCrLf & _
                            "<STMTTRNRS>" & vbCrLf & _
                            "<TRNUID>1001" & vbCrLf & _
                            "<STATUS>" & vbCrLf & _
                            "<CODE>0" & vbCrLf & _
                            "<SEVERITY>INFO" & vbCrLf & _
                            "</STATUS>" & vbCrLf & _
                            "<STMTRS>" & vbCrLf & _
                            "<CURDEF>USD" & vbCrLf & _
                            "<BANKACCTFROM>" & vbCrLf & _
                            "<BANKID>" & strBankID & vbCrLf & _
                            "<ACCTID>" & strAccID & vbCrLf & _
                            "<ACCTTYPE>CHECKING" & vbCrLf & _
                            "</BANKACCTFROM>" & vbCrLf & _
                            "<BANKTRANLIST>" & vbCrLf & _
                            "<DTSTART>" & strDataInicio & vbCrLf & _
                            "<DTEND>" & strDataFim & vbCrLf
    .Update
End With


        
'insere os dados das transações
With rsttmpDados
    .MoveFirst
    Do While Not .EOF
        rstExport.MoveFirst
        rstExport.Edit
        rstExport.Fields!TxtAExportar = rstExport.Fields!TxtAExportar & vbCrLf & _
            "<STMTTRN>" & vbCrLf & _
            .Fields!TipoTransac & vbCrLf & _
            .Fields!DataTransac & vbCrLf & _
            .Fields!Valor & vbCrLf & _
            .Fields!Fitid & vbCrLf & _
            .Fields!Memo & vbCrLf & _
            "</STMTTRN>" & vbCrLf
        rstExport.Update
        .MoveNext
    Loop
End With
    
With rstExport
    'insere o rodapé WPL
    .MoveFirst
    .Edit
    .Fields!TxtAExportar = .Fields!TxtAExportar & vbCrLf & _
        "</BANKTRANLIST>" & vbCrLf & _
        "<LEDGERBAL>" & vbCrLf & _
        "<BALAMT>" & strSaldoFinal & vbCrLf & _
        "<DTASOF>" & strDataFim & vbCrLf & _
        "</LEDGERBAL>" & vbCrLf & _
        "</STMTRS>" & vbCrLf & _
        "</STMTTRNRS>" & vbCrLf & _
        "</BANKMSGSRSV1>" & vbCrLf & _
        "</OFX>"
    .Update
End With

rstExport.Close
rstDadosOriginais.Close
rsttmpDados.Close

Set rstExport = Nothing
Set rstDadosOriginais = Nothing
Set rsttmpDados = Nothing
    
Set dbs = Nothing
    
    'crio o nome do .txt a exportar
    oldFileName = strPath & strOFXName & ".txt"
            
    'crio com sufixo .ofx o nome do arquivo para renomear o arquivo exportado com sufixo .txt
    newFileName = strPath & strOFXName & ".ofx"
            
    'exporta .txt
    DoCmd.TransferText acExportDelim, "specTxtAExportar", "TxtAExportar", oldFileName
    
    'Checa se o arquivo com o mesmo nome jé existe
    If FileOrDirExists(newFileName) = True Then
        'Se já existir, deleta
        Kill newFileName
    End If
    
    'Renomeio .txt para .ofx
    Name oldFileName As newFileName
    
'avisa que rotina foiexecutada a contento
Debug.Print MsgBox(strOFXName & " exportado com sucesso", vbOKOnly)

Exit_Function:
    Exit Function

Error_Handler:
    'erro 3021: no loop cujo fim é EOF eu comparo a data da transação do registro anterior com a data
    'da transação onde está o ponteiro. Quando chega em EOF, não ha mais data no ponteiro para comparar com a anterior
    'Assim, este erro significa apenas que o loop acabou e devolvo a rotina para o ponto onde deve continuar
    If Err.Number = 3021 Then
        GoTo Prossegue
    Else
        MsgBox Err.Description & Err.Number
        Resume Exit_Function
    End If

End Function

Function FileOrDirExists(newFileName) As Boolean
     'Macro Purpose: Function returns TRUE if the specified file
     '               or folder exists, false if not.
     'PathName     : Supports Windows mapped drives or UNC
     '             : Supports Macintosh paths
     'File usage   : Provide full file path and extension
     'Folder usage : Provide full folder path
     '               Accepts with/without trailing "\" (Windows)
     '               Accepts with/without trailing ":" (Macintosh)
     
    Dim iTemp As Integer
     
     'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(newFileName)
     
     'Check if error exists and set response appropriately
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select
     
     'Resume error checking
    On Error GoTo 0
End Function
