Sub Worksheet_Change(ByVal Target As Range)
    
    On Error GoTo ErrorHandler
    
    If Target.Address = "$C$8" And ActiveSheet.Cells(8, 3).Value <> "" Then
        Application.ScreenUpdating = False
        Dim rng As Range
        Dim numeroGaiola As String
        Dim localOrigem As String
        Dim localDestino As String
        Dim numeroColuna As Long
        
        Sheets("macro").Activate
    
        localOrigem = ActiveSheet.Cells(4, 3).Value
        localDestino = ActiveSheet.Cells(6, 3).Value
        numeroGaiola = ActiveSheet.Cells(8, 3).Value
        
        If localOrigem = localDestino Then
            
            MsgBox "A origem e destino não podem ser iguais"
            Sheets("macro").Activate
            ActiveSheet.Cells(8, 3).ClearContents
            ActiveSheet.Cells(8, 3).Select
            
            ActiveSheet.Cells(3, 5).Interior.ColorIndex = 3
            ActiveSheet.Cells(3, 5).Value = "A gaiola " & numeroGaiola & " não foi movimentada!"
            ActiveSheet.Cells(3, 5).VerticalAlignment = xlCenter
            ActiveSheet.Cells(3, 5).HorizontalAlignment = xlCenter
            
            Application.ScreenUpdating = True
                  
        Else
        
            Sheets("BD").Activate
            
            Set rng = ActiveSheet.Columns("A:A").Find(What:=numeroGaiola, _
            LookIn:=xlFormulas, LookAt:=xlWhole, SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
            
            numeroColuna = rng.Row
            
            If ActiveSheet.Cells(numeroColuna, 3).Value <> localOrigem Then
        
                MsgBox "A Gaiola não está na origem selecionada"
                Sheets("macro").Activate
                ActiveSheet.Cells(8, 3).ClearContents
                
                ActiveSheet.Cells(3, 5).Interior.ColorIndex = 3
                ActiveSheet.Cells(3, 5).Value = "A gaiola " & numeroGaiola & " não foi movimentada!"
                ActiveSheet.Cells(3, 5).VerticalAlignment = xlCenter
                ActiveSheet.Cells(3, 5).HorizontalAlignment = xlCenter
                
                Application.ScreenUpdating = True
        
            ElseIf ActiveSheet.Cells(numeroColuna, 3).Value = localOrigem Then
                ActiveSheet.Cells(numeroColuna, 3).Value = localDestino
                               
                If localOrigem = "3 - CD" Then
                    ActiveSheet.Cells(numeroColuna, 2).Value = Now
                    ActiveSheet.Cells(numeroColuna, 4).ClearContents
                    
                    Sheets("ocorrencias").Activate
                    ActiveSheet.Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
                    ActiveCell.Value = Now
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = "Envio"
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = numeroGaiola
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = localOrigem
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = localDestino
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = Environ$("computername")
                    
                ElseIf localDestino = "3 - CD" Then
                    ActiveSheet.Cells(numeroColuna, 4).Value = Now
                    
                    Sheets("ocorrencias").Activate
                    ActiveSheet.Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
                    ActiveCell.Value = Now
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = "Retorno"
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = numeroGaiola
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = localOrigem
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = localDestino
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = Environ$("computername")
                    
            
                Else
                    MsgBox "Transferência direta de uma loja para outra"
                    ActiveSheet.Cells(numeroColuna, 2).Value = Now
                    ActiveSheet.Cells(numeroColuna, 4).ClearContents
                    
                    Sheets("ocorrencias").Activate
                    ActiveSheet.Range("A" & Rows.Count).End(xlUp).Offset(1, 0).Select
                    ActiveCell.Value = Now
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = "Envio Direto"
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = numeroGaiola
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = localOrigem
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = localDestino
                    ActiveCell.Offset(0, 1).Select
                    ActiveCell.Value = Environ$("computername")

                End If

                Sheets("macro").Activate
                ActiveSheet.Cells(8, 3).ClearContents
                ActiveSheet.Cells(3, 5).Interior.ColorIndex = 4
                ActiveSheet.Cells(3, 5).Value = "Gaiola " & numeroGaiola & " movimentada com sucesso!"
                ActiveSheet.Cells(3, 5).VerticalAlignment = xlCenter
                ActiveSheet.Cells(3, 5).HorizontalAlignment = xlCenter

            End If
        End If
    End If
    ActiveSheet.Cells(8, 3).Select
    ActiveWorkbook.RefreshAll
    Application.ScreenUpdating = True
    
Exit Sub
ErrorHandler:
    MsgBox "A Gaiola informada não existe"
    Sheets("macro").Activate
    ActiveSheet.Cells(8, 3).ClearContents
    ActiveWorkbook.RefreshAll
    ActiveSheet.Cells(8, 3).Select
    
    ActiveSheet.Cells(3, 5).Interior.ColorIndex = 3
    ActiveSheet.Cells(3, 5).Value = "A gaiola " & numeroGaiola & " não existe!"
    ActiveSheet.Cells(3, 5).VerticalAlignment = xlCenter
    ActiveSheet.Cells(3, 5).HorizontalAlignment = xlCenter
    
    Application.ScreenUpdating = True
    Exit Sub
End Sub