Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo Fim
    Application.EnableEvents = False

    ' GERAR ESCALA AO MUDAR DATA
    If Not Intersect(Target, Me.Range("C1")) Is Nothing Then
        
        If Target.Value <> "" Then
            
            If MsgBox("Gerar nova escala?", vbYesNo + vbQuestion) = vbYes Then
                Call GerarEscalaSemanal
            End If
            
        End If
        
    End If

    ' PROCESSAR FALTA
    If Not Intersect(Target, Me.Range("E4:E100")) Is Nothing Then
        
        If Target.Value <> "" Then
            Call ProcessarFalta(Target.Row)
        End If
        
    End If

Fim:
    Application.EnableEvents = True

End Sub