Option Explicit

Const LIMITE_CICLO As Integer = 2

' =========================================
' GERAR ESCALA
' =========================================
Sub GerarEscalaSemanal()

    Application.ScreenUpdating = False
    
    Dim wsEscala As Worksheet
    Dim wsHistorico As Worksheet
    Dim wsFeriados As Worksheet
    
    Set wsEscala = ThisWorkbook.Sheets("ESCALA")
    Set wsHistorico = ThisWorkbook.Sheets("HISTORICO")
    Set wsFeriados = ThisWorkbook.Sheets("FERIADOS + FOLGAS")
    
    Dim dataInicio As Date
    dataInicio = wsEscala.Range("C1").Value
    
    If dataInicio = 0 Then Exit Sub
    
    ' aqui ele ajusta para  segunda feira, atualiza automatico
    Dim diaSemana As Integer
    diaSemana = Weekday(dataInicio, vbMonday)
    
    If diaSemana > 1 Then
        dataInicio = dataInicio - (diaSemana - 1)
    End If
    
    wsEscala.Range("A4:F100").ClearContents
    
    ' =============================
    ' COLABORADORES DINÂMICOS
    ' =============================
    Dim ultimaLinha As Long
    ultimaLinha = wsHistorico.Cells(wsHistorico.Rows.Count, 1).End(xlUp).Row
    
    Dim totalColab As Integer
    totalColab = ultimaLinha - 1
    
    If totalColab < 2 Then
        MsgBox "É necessário ter pelo menos 2 colaboradores!", vbExclamation
        Exit Sub
    End If
    
    Dim colaboradores() As Variant
    ReDim colaboradores(1 To totalColab, 1 To 4)
    
    Dim i As Integer
    
    For i = 1 To totalColab
        colaboradores(i, 1) = wsHistorico.Cells(i + 1, 1).Value  ' NOME
        colaboradores(i, 2) = wsHistorico.Cells(i + 1, 2).Value  ' QTD_MES
        colaboradores(i, 3) = wsHistorico.Cells(i + 1, 3).Value  ' QTD_GERAL
        colaboradores(i, 4) = wsHistorico.Cells(i + 1, 4).Value  ' QTD_CICLO
    Next i
    
    ' =============================
    ' VERIFICAR E RESETAR CICLO
    ' =============================
    Dim resetouCiclo As Boolean
    resetouCiclo = False
    
    If DeveResetarCiclo(wsHistorico, totalColab) Then
        Call ResetarCiclo(wsHistorico, totalColab)
        resetouCiclo = True
        
        ' atualiza o array e depois reseta ele
        For i = 1 To totalColab
            colaboradores(i, 4) = 0  ' Zera a QTD_CICLO no array apos atingit a quantidade
        Next i
        
        ' embaralha os nomes
        Call Embaralhar(colaboradores, totalColab)
        
        MsgBox "Ciclo completo! Duplas embaralhadas para nova rotação.", vbInformation
    End If
    
    ' =============================
    ' ORDENAR 
    ' =============================
    If Not resetouCiclo Then
        Call OrdenarColaboradores(colaboradores, totalColab)
    End If
    
    ' =============================
    ' GERAR ESCALA
    ' =============================
    Dim linha As Integer: linha = 4
    Dim indice As Integer: indice = 1
    Dim dataAtual As Date
    Dim k As Long, ehFeriado As Boolean
    
    For i = 0 To 4  ' Segunda a Sexta
        
        dataAtual = dataInicio + i
        ehFeriado = False
        
        ' Verificar se é feriado
        For k = 2 To wsFeriados.Cells(wsFeriados.Rows.Count, 1).End(xlUp).Row
            If Int(wsFeriados.Cells(k, 1).Value) = Int(dataAtual) Then
                ehFeriado = True
                Exit For
            End If
        Next k
        
        wsEscala.Cells(linha, 1).Value = dataAtual
        wsEscala.Cells(linha, 1).NumberFormat = "dd/mm/yyyy"
        
        If ehFeriado Then
            
            wsEscala.Cells(linha, 4).Value = "BLOQUEADO"
            wsEscala.Cells(linha, 6).Value = wsFeriados.Cells(k, 3).Value
            
        Else
            
            Dim aux1 As String, aux2 As String
            
            ' Validar índice
            If indice > totalColab Then indice = indice - totalColab
            
            aux1 = colaboradores(indice, 1)
            
            If indice + 1 > totalColab Then
                aux2 = colaboradores(1, 1)
            Else
                aux2 = colaboradores(indice + 1, 1)
            End If
            
            wsEscala.Cells(linha, 2).Value = aux1
            wsEscala.Cells(linha, 3).Value = aux2
            wsEscala.Cells(linha, 4).Value = aux1 & " + " & aux2
            
            indice = indice + 2
            
        End If
        
        linha = linha + 1
        
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Escala gerada com sucesso!", vbInformation

End Sub

' =========================================
' ORDENAR COLABORADORES
' =========================================
Sub OrdenarColaboradores(colaboradores As Variant, totalColab As Integer)
    
    Dim i As Integer, j As Integer
    Dim temp1, temp2, temp3, temp4
    Dim swap As Boolean
    
    ' Bubble Sort por: QTD_MES -> QTD_GERAL -> NOME (esse aqui e so pra caso uma das alternativas nao dar certo, pra nao quebrar kkk)
    For i = 1 To totalColab - 1
        For j = i + 1 To totalColab
            
            swap = False
            
            ' Primeiro critério: menor QTD_MES
            If colaboradores(j, 2) < colaboradores(i, 2) Then
                swap = True
                
            ElseIf colaboradores(j, 2) = colaboradores(i, 2) Then
                
                ' Segundo critério: menor QTD_GERAL
                If colaboradores(j, 3) < colaboradores(i, 3) Then
                    swap = True
                    
                ElseIf colaboradores(j, 3) = colaboradores(i, 3) Then
                    
                    ' Terceiro critério: ordem alfabética
                    If colaboradores(j, 1) < colaboradores(i, 1) Then
                        swap = True
                    End If
                    
                End If
                
            End If
            
            If swap Then
                
                temp1 = colaboradores(i, 1)
                temp2 = colaboradores(i, 2)
                temp3 = colaboradores(i, 3)
                temp4 = colaboradores(i, 4)
                
                colaboradores(i, 1) = colaboradores(j, 1)
                colaboradores(i, 2) = colaboradores(j, 2)
                colaboradores(i, 3) = colaboradores(j, 3)
                colaboradores(i, 4) = colaboradores(j, 4)
                
                colaboradores(j, 1) = temp1
                colaboradores(j, 2) = temp2
                colaboradores(j, 3) = temp3
                colaboradores(j, 4) = temp4
                
            End If
            
        Next j
    Next i

End Sub

' =========================================
' PROCESSAR FALTA
' =========================================
Sub ProcessarFalta(linhaFalta As Integer)

    Dim wsEscala As Worksheet
    Dim wsHistorico As Worksheet
    
    Set wsEscala = ThisWorkbook.Sheets("ESCALA")
    Set wsHistorico = ThisWorkbook.Sheets("HISTORICO")
    
    Dim faltoso As String
    faltoso = wsEscala.Cells(linhaFalta, 5).Value
    
    If faltoso = "" Then Exit Sub
    
    ' =============================
    ' IDENTIFICAR DUPLA
    ' =============================
    Dim aux1 As String, aux2 As String
    aux1 = wsEscala.Cells(linhaFalta, 2).Value
    aux2 = wsEscala.Cells(linhaFalta, 3).Value
    
    Dim presente As String
    
    If faltoso = aux1 Then
        presente = aux2
    ElseIf faltoso = aux2 Then
        presente = aux1
    Else
        MsgBox "Nome da falta não corresponde à dupla.", vbExclamation
        Exit Sub
    End If
    
    ' =============================
    ' PROCURAR PRÓXIMO DIA VÁLIDO
    ' =============================
    Dim proximaLinha As Integer
    proximaLinha = linhaFalta + 1
    
    Do While proximaLinha <= 100
        
        If wsEscala.Cells(proximaLinha, 4).Value <> "BLOQUEADO" _
           And wsEscala.Cells(proximaLinha, 4).Value <> "" Then
            Exit Do
        End If
        
        proximaLinha = proximaLinha + 1
        
    Loop
    
    If proximaLinha > 100 Then
        MsgBox "Não há dia disponível para substituição.", vbExclamation
        Exit Sub
    End If
    
    ' =============================
    ' PEGAR DUPLA DO PRÓXIMO DIA
    ' =============================
    Dim proximoAux1 As String, proximoAux2 As String
    proximoAux1 = wsEscala.Cells(proximaLinha, 2).Value
    proximoAux2 = wsEscala.Cells(proximaLinha, 3).Value
    
    ' =============================
    ' ESCOLHER SUBSTITUTO (MAIS JUSTO)
    ' =============================
    Dim cont1 As Long, cont2 As Long
    Dim linhaHist1 As Variant, linhaHist2 As Variant
    
    linhaHist1 = Application.Match(proximoAux1, wsHistorico.Range("A:A"), 0)
    linhaHist2 = Application.Match(proximoAux2, wsHistorico.Range("A:A"), 0)
    
    If IsError(linhaHist1) Or IsError(linhaHist2) Then
        MsgBox "Colaborador não encontrado no histórico.", vbExclamation
        Exit Sub
    End If
    
    cont1 = wsHistorico.Cells(linhaHist1, 4).Value  ' QTD_CICLO
    cont2 = wsHistorico.Cells(linhaHist2, 4).Value  ' QTD_CICLO
    
    Dim substituto As String, ficaNaProxima As String
    
    If cont1 < cont2 Then
        substituto = proximoAux1
        ficaNaProxima = proximoAux2
        
    ElseIf cont2 < cont1 Then
        substituto = proximoAux2
        ficaNaProxima = proximoAux1
        
    Else
        ' Desempate alfabético
        If proximoAux1 < proximoAux2 Then
            substituto = proximoAux1
            ficaNaProxima = proximoAux2
        Else
            substituto = proximoAux2
            ficaNaProxima = proximoAux1
        End If
    End If
    
    ' =============================
    ' REORGANIZAR ESCALA
    ' =============================
    
    ' Dia da falta
    wsEscala.Cells(linhaFalta, 2).Value = presente
    wsEscala.Cells(linhaFalta, 3).Value = substituto
    wsEscala.Cells(linhaFalta, 4).Value = presente & " + " & substituto
    
    ' Próximo dia (reposição)
    wsEscala.Cells(proximaLinha, 2).Value = faltoso
    wsEscala.Cells(proximaLinha, 3).Value = ficaNaProxima
    wsEscala.Cells(proximaLinha, 4).Value = faltoso & " + " & ficaNaProxima
    
    ' =============================
    ' OBSERVAÇÃO
    ' =============================
    wsEscala.Cells(linhaFalta, 6).Value = _
        "FALTA: " & faltoso & " | SUB: " & substituto
    
    wsEscala.Cells(proximaLinha, 6).Value = _
        "REPOSIÇÃO: " & faltoso
    
    ' Limpar célula de falta após processamento
    wsEscala.Cells(linhaFalta, 5).ClearContents
    
    MsgBox "Substituição realizada com sucesso!", vbInformation

End Sub

' =========================================
' FINALIZAR SEMANA
' =========================================
Sub FinalizarSemana()

    Dim wsEscala As Worksheet
    Dim wsHistorico As Worksheet
    
    Set wsEscala = ThisWorkbook.Sheets("ESCALA")
    Set wsHistorico = ThisWorkbook.Sheets("HISTORICO")
    
    Dim i As Integer
    Dim totalDias As Integer
    totalDias = 0
    
    For i = 4 To 8  ' Segunda a Sexta
        
        If wsEscala.Cells(i, 4).Value <> "BLOQUEADO" _
           And wsEscala.Cells(i, 4).Value <> "" Then
            
            Call AtualizarContador(wsHistorico, wsEscala.Cells(i, 2).Value)
            Call AtualizarContador(wsHistorico, wsEscala.Cells(i, 3).Value)
            
            totalDias = totalDias + 1
            
        End If
        
    Next i
    
    MsgBox "Semana finalizada! " & totalDias & " dias processados.", vbInformation

End Sub

' =========================================
' FECHAR MÊS
' =========================================
Sub FecharMes()

    Dim wsHistorico As Worksheet
    Dim wsIndicador As Worksheet
    Dim wsEscala As Worksheet
    
    Set wsHistorico = ThisWorkbook.Sheets("HISTORICO")
    Set wsIndicador = ThisWorkbook.Sheets("INDICADORES")
    Set wsEscala = ThisWorkbook.Sheets("ESCALA")
    
    Dim mesReferencia As String
    mesReferencia = Format(wsEscala.Range("C1").Value, "mm/yyyy")
    
    Dim ultimaColuna As Long
    ultimaColuna = wsIndicador.Cells(1, wsIndicador.Columns.Count).End(xlToLeft).Column + 1
    
    wsIndicador.Cells(1, ultimaColuna).Value = mesReferencia
    
    Dim ultimaLinha As Long
    ultimaLinha = wsHistorico.Cells(wsHistorico.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Integer
    
    For i = 2 To ultimaLinha
        
        ' Copiar nome e quantidade do mês
        wsIndicador.Cells(i, 1).Value = wsHistorico.Cells(i, 1).Value
        wsIndicador.Cells(i, ultimaColuna).Value = wsHistorico.Cells(i, 2).Value
        
        ' Zerar contador mensal
        wsHistorico.Cells(i, 2).Value = 0
        
    Next i
    
    MsgBox "Mês " & mesReferencia & " fechado com sucesso!", vbInformation

End Sub

' =========================================
' ATUALIZAR CONTADORES
' =========================================
Sub AtualizarContador(ws As Worksheet, nome As String)

    Dim linha As Variant
    
    linha = Application.Match(nome, ws.Range("A:A"), 0)
    
    If Not IsError(linha) Then
        
        ws.Cells(linha, 2).Value = ws.Cells(linha, 2).Value + 1  ' QTD_MES
        ws.Cells(linha, 3).Value = ws.Cells(linha, 3).Value + 1  ' QTD_GERAL
        ws.Cells(linha, 4).Value = ws.Cells(linha, 4).Value + 1  ' QTD_CICLO
        
    End If

End Sub

' =========================================
' VERIFICAR SE DEVE RESETAR CICLO
' =========================================
Function DeveResetarCiclo(wsHistorico As Worksheet, totalColab As Integer) As Boolean

    Dim i As Integer
    Dim minCiclo As Long
    
    minCiclo = 999999
    
    ' Encontrar o menor valor de QTD_CICLO
    For i = 2 To totalColab + 1
        If wsHistorico.Cells(i, 4).Value < minCiclo Then
            minCiclo = wsHistorico.Cells(i, 4).Value
        End If
    Next i
    
    DeveResetarCiclo = (minCiclo >= LIMITE_CICLO)

End Function

' =========================================
' RESETAR CICLO
' =========================================
Sub ResetarCiclo(wsHistorico As Worksheet, totalColab As Integer)

    Dim i As Integer
    
    For i = 2 To totalColab + 1
        wsHistorico.Cells(i, 4).Value = 0  
    Next i

End Sub

' =========================================
' EMBARALHAR (Fisher-Yates Shuffle)
' =========================================
Sub Embaralhar(colaboradores As Variant, totalColab As Integer)

    Dim i As Integer, j As Integer
    Dim t1, t2, t3, t4
    
    Randomize Timer 
    
    ' Algoritmo Fisher-Yates
    For i = totalColab To 2 Step -1
        
        j = Int(Rnd * i) + 1  
        
        
        t1 = colaboradores(i, 1)
        t2 = colaboradores(i, 2)
        t3 = colaboradores(i, 3)
        t4 = colaboradores(i, 4)
        
        colaboradores(i, 1) = colaboradores(j, 1)
        colaboradores(i, 2) = colaboradores(j, 2)
        colaboradores(i, 3) = colaboradores(j, 3)
        colaboradores(i, 4) = colaboradores(j, 4)
        
        colaboradores(j, 1) = t1
        colaboradores(j, 2) = t2
        colaboradores(j, 3) = t3
        colaboradores(j, 4) = t4
        
    Next i

End Sub

' =========================================
' FUNÇÃO AUXILIAR: EXIBIR STATUS DO CICLO
' =========================================
Sub ExibirStatusCiclo()

    Dim wsHistorico As Worksheet
    Set wsHistorico = ThisWorkbook.Sheets("HISTORICO")
    
    Dim ultimaLinha As Long
    ultimaLinha = wsHistorico.Cells(wsHistorico.Rows.Count, 1).End(xlUp).Row
    
    Dim mensagem As String
    mensagem = "STATUS DO CICLO:" & vbCrLf & vbCrLf
    
    Dim i As Integer
    For i = 2 To ultimaLinha
        mensagem = mensagem & wsHistorico.Cells(i, 1).Value & ": " & _
                   wsHistorico.Cells(i, 4).Value & "/" & LIMITE_CICLO & vbCrLf
    Next i
    
    MsgBox mensagem, vbInformation, "Ciclo Atual"

End Sub