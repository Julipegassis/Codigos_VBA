Attribute VB_Name = "EXPT_ASSIS"
' Removido no dia 21/02/2023 funcoes utilizadas em outras planilhas.


Function RetornaMedData(Antigo As Date, Atual As Date) As Date
    If ((Antigo + Atual / 2) < (Antigo * 0.3)) Or ((Antigo + Atual / 2) < (Atual * 0.3)) Then
        RetornaMedData = (#11:59:00 PM# + #12:01:00 AM# + Atual + Antigo) / 2
        
      Else
            RetornaMedData = (Antigo + Atual) / 2
        End If

End Function
Function RetornaDobra(INICIO As Date, FIM As Date) As Boolean
RetornaDobra = False
If INICIO > FIM Then
    RetornaDobra = True
    End If

End Function
Function RetornaDifData(INICIO As Date, FIM As Date) As Date
    If INICIO > FIM Then
        RetornaDifData = #11:59:00 PM# - INICIO
        RetornaDifData = RetornaDifData + FIM + #12:01:00 AM#
      Else
            RetornaDifData = FIM - INICIO
        End If

End Function


Function Hora50(FERIADO As Boolean, Proximodia As Boolean, INICIO As Date, FIM As Date, tinicio As Date, tfim As Date) As Date
Dim Tot As Date
Dim escolhaf As Integer
Dim escolhat As Integer

Tot = RetornaDifData(INICIO, FIM)
If FERIADO = True And Proximodia = True Then
    escolhaf = 1 'feriado e proximo dia feriado
    End If
If FERIADO = False And Proximodia = True Then
    escolhaf = 2 ' dia normal e proximo dia feriado
    End If
If FERIADO = True And Proximodia = False Then
    escolhaf = 3 ' feriado e proximo dia normal
    End If
If FERIADO = False And Proximodia = False Then
    escolhaf = 4 ' dia normal e proximo dia normal
    End If
    
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 1 ' funcionario noturno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 2 ' funcionario noturno e saiu diurno
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 3 ' funcionario é diurno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 4 ' funcionario é diurno e saiu diurno
    End If

Select Case escolhaf
Case 1 'feriado e proximo dia feriado
    Select Case escolhat
        Case 1 To 4
        Hora50 = #12:00:00 AM#
        End Select
        
Case 2 ' dia normal e proximo dia feriado
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            
                Hora50 = #12:00:00 AM#
                        
        Case 2 ' noturno e saiu diurno
                Hora50 = #12:00:00 AM#
        Case 3 ' diurno e saiu noturno
                If tfim >= #10:00:00 PM# Then
                'se a hora extra comeca depois de 22:00 ele não tem 50%, mas pode receber
                
                    Hora50 = #12:00:00 AM#
                    End If
                
                If tfim < #10:00:00 PM# And INICIO < #10:00:00 PM# Then
                    If INICIO <= tfim Then
                              Hora50 = RetornaDifData(tfim, #10:00:00 PM#)
                            End If
                    If INICIO > tfim Then
                              Hora50 = RetornaDifData(INICIO, #10:00:00 PM#)
                            End If
                    End If
                                
        Case 4 ' diurno e saiu diurno
                If tfim > #10:00:00 PM# Then
                        Hora50 = #12:00:00 AM#
                Else
                If FIM > tfim And FIM > #10:00:00 PM# Then
                    If INICIO < tfim Then
                        Hora50 = RetornaDifData(tfim, #10:00:00 PM#)
                        End If
                    If INICIO >= tfim Then
                        Hora50 = RetornaDifData(INICIO, #10:00:00 PM#)
                        End If
                    End If
                If FIM > tfim And FIM < #10:00:00 PM# Then
                    If INICIO < tfim Then
                            Hora50 = RetornaDifData(tfim, FIM)
                            End If
                        If INICIO >= tfim Then
                            Hora50 = RetornaDifData(INICIO, FIM)
                            End If
                     End If
                End If
        End Select

'case 3 para 50% diurna

Case 3
'feriado e proximo dia normal
If INICIO > FIM Then
    If FIM >= #5:00:00 AM# Then
        Hora50 = FIM - #5:00:00 AM#
    Else
        Hora50 = 0
    End If

End If

Case 4 ' dia normal e proximo dia normal
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            If FIM < #5:00:00 AM# Then
                Hora50 = #12:00:00 AM#
                End If
            If FIM > #5:00:00 AM# And tfim < #5:00:00 AM# Then
                Hora50 = RetornaDifData(#5:00:00 AM#, FIM)
                End If
            If FIM > #5:00:00 AM# And tfim >= #5:00:00 AM# Then
            
                Hora50 = RetornaDifData(tfim, FIM)
                End If
            
        Case 2 ' noturno e saiu diurno
            If FIM < #10:00:00 AM# And INICIO < #10:00:00 AM# Then
                
                    If FIM > #5:00:00 AM# And FIM > tfim Then
                        If INICIO <= #5:00:00 AM# Then
                                If tfim < #5:00:00 AM# Then
                                    Hora50 = RetornaDifData(#5:00:00 AM#, FIM)
                                    Else
                                    Hora50 = RetornaDifData(tfim, FIM)
                                End If
                        Else
                            If tfim < #5:00:00 AM# Then
                                Hora50 = RetornaDifData(INICIO, FIM)
                            Else
                                       If INICIO < tfim Then
                                            Hora50 = RetornaDifData(tfim, FIM)
                                        Else
                                            Hora50 = RetornaDifData(INICIO, FIM)
                                       End If
                            End If
                        End If
                   Else
                    Hora50 = #12:00:00 AM#
                   End If
              Else
                    Hora50 = #12:00:00 AM#
              End If
        Case 3 ' diurno e saiu noturno
                If tfim >= #10:00:00 PM# And FIM > #5:00:00 AM# Then
                'se a hora extra comeca depois de 22:00 ele não tem 50%, mas pode receber
                ' se ele passar de 05:00, em se tratando de funcionario diurno ja sera extra
                    Hora50 = RetornaDifData(#5:00:00 AM#, FIM)
                    End If
                If tfim < #10:00:00 PM# And INICIO < #10:00:00 PM# And FIM > #5:00:00 AM# Then
                'se a jornada termina antes de 22:00 tem direito a hora extra 50% se ele iniciou antes das 22:00
                    If INICIO <= tfim Then
                        'e como passou de 05:00 ele tem que receber mais 50%
                            Hora50 = RetornaDifData(#5:00:00 AM#, FIM) + RetornaDifData(tfim, #10:00:00 PM#)
                            End If
                    If INICIO > tfim Then
                        'e como passou de 05:00 ele tem que receber mais 50%
                            Hora50 = RetornaDifData(#5:00:00 AM#, FIM) + RetornaDifData(INICIO, #10:00:00 PM#)
                            End If
                            
                    If INICIO <= tfim Then
                        'e como passou de 05:00 ele tem que receber mais 50%
                            Hora50 = RetornaDifData(#5:00:00 AM#, FIM) + RetornaDifData(INICIO, #10:00:00 PM#)
                            End If
                    End If
                If tfim < #10:00:00 PM# And INICIO < #10:00:00 PM# And FIM <= #5:00:00 AM# Then
                    If INICIO <= tfim Then
                              Hora50 = RetornaDifData(tfim, #10:00:00 PM#)
                            End If
                    If INICIO > tfim Then
                              Hora50 = RetornaDifData(INICIO, #10:00:00 PM#)
                            End If
                    End If
                If tfim > #10:00:00 PM# And FIM > #5:00:00 AM# Then
                    Hora50 = RetornaDifData(#5:00:00 AM#, FIM)
                    End If
                
        Case 4 ' diurno e saiu diurno
                If tfim > #10:00:00 PM# Then
                        Hora50 = #12:00:00 AM#
                Else
                If FIM > tfim And FIM > #10:00:00 PM# Then
                    If INICIO < tfim Then
                        Hora50 = RetornaDifData(tfim, #10:00:00 PM#)
                        End If
                    If INICIO >= tfim Then
                        Hora50 = RetornaDifData(INICIO, #10:00:00 PM#)
                        End If
                    End If
                If FIM > tfim And FIM <= #10:00:00 PM# Then
                    If INICIO < tfim Then
                            Hora50 = RetornaDifData(tfim, FIM)
                            End If
                        If INICIO >= tfim Then
                            Hora50 = RetornaDifData(INICIO, FIM)
                            End If
                     End If
                End If
        End Select
End Select
If Hora50 > Tot Then
    MsgBox ("Retorno incorreto, consulte administrador")
    End If
End Function
Function HoraNormal(FERIADO As Boolean, Proximodia As Boolean, INICIO As Date, FIM As Date, tinicio As Date, tfim As Date) As Date
Dim Tot As Date
Dim escolhaf As Integer
Dim escolhat As Integer

Tot = RetornaDifData(INICIO, FIM)
If FERIADO = True And Proximodia = True Then
    escolhaf = 1 'feriado e proximo dia feriado
    End If
If FERIADO = False And Proximodia = True Then
    escolhaf = 2 ' dia normal e proximo dia feriado
    End If
If FERIADO = True And Proximodia = False Then
    escolhaf = 3 ' feriado e proximo dia normal
    End If
If FERIADO = False And Proximodia = False Then
    escolhaf = 4 ' dia normal e proximo dia normal
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 1 ' funcionario noturno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 2 ' funcionario noturno e saiu diurno
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 3 ' funcionario é diurno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 4 ' funcionario é diurno e saiu diurno
    End If

Select Case escolhaf
Case 1 'feriado e proximo dia feriado
    Select Case escolhat
        Case 1 To 4
        HoraNormal = #12:00:00 AM#
        End Select
        
Case 2 ' dia normal e proximo dia feriado
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            If FIM > tfim And INICIO < tfim Then
                    HoraNormal = RetornaDifData(INICIO, tfim)
                    Else
                    HoraNormal = RetornaDifData(INICIO, FIM)
             End If
                        
        Case 2 ' noturno e saiu diurno
            If INICIO < #10:00:00 AM# And FIM < #10:00:00 AM# Then
                    If FIM > tfim And INICIO < tfim Then
                        HoraNormal = RetornaDifData(INICIO, tfim)
                        Else
                        If INICIO < tfim Then
                        HoraNormal = RetornaDifData(INICIO, FIM)
                        End If
                    End If
                Else
                    If INICIO < tfim Then
                    HoraNormal = RetornaDifData(INICIO, FIM)
                    End If
                End If
        Case 3 ' diurno e saiu noturno
                If INICIO < tfim Then
                    HoraNormal = RetornaDifData(INICIO, tfim)
                   End If
        Case 4 ' diurno e saiu diurno
                If INICIO < tfim Then
                    If FIM > tfim Then
                        HoraNormal = RetornaDifData(INICIO, tfim)
                        Else
                        HoraNormal = RetornaDifData(INICIO, FIM)
                    End If
                End If
                
        End Select

Case 3 'feriado e proximo dia normal
    Select Case escolhat
        Case 1 To 4
        HoraNormal = #12:00:00 AM#
        End Select

Case 4 ' dia normal e proximo dia normal
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            If FIM > tfim Then
                HoraNormal = RetornaDifData(INICIO, tfim)
                Else
                HoraNormal = RetornaDifData(INICIO, FIM)
                End If
            
        Case 2 ' noturno e saiu diurno
            If FIM < #10:00:00 AM# And INICIO < #10:00:00 AM# Then
                
                    If FIM > tfim Then
                        If INICIO < tfim Then
                            HoraNormal = RetornaDifData(INICIO, tfim)
                        End If
                    Else
                            HoraNormal = RetornaDifData(INICIO, FIM)
                    End If
                Else
                  HoraNormal = RetornaDifData(INICIO, FIM)
                End If
        Case 3 ' diurno e saiu noturno
                If INICIO < tfim Then
                    HoraNormal = RetornaDifData(INICIO, tfim)
                    End If
                
                
        Case 4 ' diurno e saiu diurno
                If INICIO < tfim Then
                    If FIM > tfim Then
                        HoraNormal = RetornaDifData(INICIO, tfim)
                    Else
                    HoraNormal = RetornaDifData(INICIO, FIM)
                    End If
                End If
                
        End Select
End Select
If HoraNormal > Tot Then
    MsgBox ("Retorno incorreto, consulte administrador")
    End If
End Function

Function HoraAdNot(INICIO As Date, FIM As Date) As Date
If INICIO > FIM Then
    '===========================FUNCIONARIO SAIU A NOITE
        If INICIO <= #10:00:00 PM# Then
            If FIM > #5:00:00 AM# Then
                HoraAdNot = #7:00:00 AM#
            Else
         
            HoraAdNot = RetornaDifData(#12:00:00 AM#, FIM) + #2:00:00 AM#
            End If
        Else
            If FIM > #5:00:00 AM# Then
            HoraAdNot = #5:00:00 AM# + RetornaDifData(INICIO, #12:00:00 AM#)
            Else
            HoraAdNot = RetornaDifData(#12:00:00 AM#, FIM) + RetornaDifData(INICIO, #12:00:00 AM#)
            End If
        End If
Else
If INICIO < FIM Then
    If FIM < #10:00:00 PM# Then
    HoraAdNot = #12:00:00 AM#
    Else
        If INICIO < #10:00:00 PM# Then
        HoraAdNot = RetornaDifData(#10:00:00 PM#, FIM)
        Else
        HoraAdNot = RetornaDifData(INICIO, FIM)
        End If
    End If
End If
End If
      
End Function

Function Hora50N(FERIADO As Boolean, Proximodia As Boolean, INICIO As Date, FIM As Date, tinicio As Date, tfim As Date) As Date
Dim Tot As Date
Dim escolhaf As Integer
Dim escolhat As Integer

Tot = RetornaDifData(INICIO, FIM)
If FERIADO = True And Proximodia = True Then
    escolhaf = 1 'feriado e proximo dia feriado
    End If
If FERIADO = False And Proximodia = True Then
    escolhaf = 2 ' dia normal e proximo dia feriado
    End If
If FERIADO = True And Proximodia = False Then
    escolhaf = 3 ' feriado e proximo dia normal
    End If
If FERIADO = False And Proximodia = False Then
    escolhaf = 4 ' dia normal e proximo dia normal
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 1 ' funcionario noturno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 2 ' funcionario noturno e saiu diurno
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 3 ' funcionario é diurno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 4 ' funcionario é diurno e saiu diurno
    End If

Select Case escolhaf
Case 1 'feriado e proximo dia feriado
    Select Case escolhat
        Case 1 To 4
        Hora50N = #12:00:00 AM#
        End Select
        
Case 2 ' dia normal e proximo dia feriado
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            
                Hora50N = #12:00:00 AM#
                        
        Case 2 ' noturno e saiu diurno
                Hora50N = #12:00:00 AM#
        Case 3 ' diurno e saiu noturno
                If tfim >= #10:00:00 PM# And INICIO < tfim Then
                'se a hora extra comeca depois de 22:00 ele não tem 50%, mas pode receber
                
                    Hora50N = RetornaDifData(tfim, #11:59:00 PM#) + #12:01:00 AM#
                    End If
                If tfim >= #10:00:00 PM# And INICIO > tfim Then
                'se a hora extra comeca depois de 22:00 ele não tem 50%, mas pode receber
                
                    Hora50N = RetornaDifData(INICIO, #11:59:00 PM#) + #12:01:00 AM#
                    End If
                
                If tfim < #10:00:00 PM# And INICIO < #10:00:00 PM# Then
                    
                              Hora50N = RetornaDifData(#10:00:00 PM#, #11:59:00 PM#) + #12:01:00 AM#
                            End If
                If tfim < #10:00:00 PM# And INICIO >= #10:00:00 PM# Then
                    
                              Hora50N = RetornaDifData(INICIO, #11:59:00 PM#) + #12:01:00 AM#
                            End If
                                
        Case 4 ' diurno e saiu diurno
                If FIM > #10:00:00 PM# And FIM > tfim Then
                'funcionario saiu mais de 22:00 e além de seu turno
                    If tfim > #10:00:00 PM# Then
                    'turno do funcionario era maior que 22:00
                        If INICIO > tfim Then
                        'inicio do trabalho foi maior que sua saída de turno
                            Hora50N = RetornaDifData(INICIO, FIM)
                            Else
                            'iniciou o trabalho antes de sua saída de turno
                            Hora50N = RetornaDifData(tfim, FIM)
                        End If
                    Else
                    'turno do funcionário é menor que 22:00
                        If INICIO >= #10:00:00 PM# Then
                            Hora50N = RetornaDifData(INICIO, FIM)
                            Else
                            Hora50N = RetornaDifData(#10:00:00 PM#, FIM)
                        End If
                    End If
                End If
                        
                      
        End Select


'case 3 para 50% noturna
Case 3 'feriado e proximo dia normal

If INICIO > FIM Then
    ' funcionario saiu madrugada e inicio depois de 10:00
        If FIM >= #5:00:00 AM# Then
            Hora50N = #5:00:00 AM#
        Else
            Hora50N = FIM
        End If
        
End If

Case 4 ' dia normal e proximo dia normal
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            If FIM <= #5:00:00 AM# And FIM > tfim Then
                Hora50N = RetornaDifData(tfim, FIM)
                End If
            If FIM > #5:00:00 AM# And tfim < #5:00:00 AM# Then
                Hora50N = RetornaDifData(tfim, #5:00:00 AM#)
                End If
            If FIM > #5:00:00 AM# And tfim >= #5:00:00 AM# Then
            
                Hora50N = #12:00:00 AM#
                End If
            
        Case 2 ' noturno e saiu diurno
                If FIM < #10:00:00 AM# And INICIO < #10:00:00 AM# Then
                    If FIM > tfim Then
                        If FIM <= #5:00:00 AM# Then
                        Hora50N = RetornaDifData(tfim, FIM)
                        Else
                        Hora50N = RetornaDifData(tfim, #5:00:00 AM#)
                        End If
                    End If
                Else
                Hora50N = #12:00:00 AM#
                End If
        Case 3 ' diurno e saiu noturno
                If INICIO > tfim Then
                    If tfim >= #10:00:00 PM# And FIM > #5:00:00 AM# Then
                        'se a hora extra comeca depois de 22:00 ele tem 50% noturno
                        ' se ele passar de 05:00, ele também tem mais 05:00
                            Hora50N = RetornaDifData(INICIO, #11:59:00 PM#) + #5:01:00 AM#
                            End If
                    If tfim < #10:00:00 PM# And FIM > #5:00:00 AM# Then
                    'se a jornada termina antes de 22:00 passa a computar depois das 22:00
                     'e como passou de 05:00 ele tem mais 05:00
                            If INICIO > #10:00:00 PM# Then
                                Hora50N = RetornaDifData(INICIO, #11:59:00 PM#) + #5:01:00 AM#
                                Else
                                Hora50N = #7:00:00 AM#
                             End If
                        End If
                    If tfim >= #10:00:00 PM# And FIM <= #5:00:00 AM# Then
                        'se a hora extra comeca depois de 22:00 ele tem 50% noturno
                        ' se ele passar de 05:00, ele também tem mais 05:00
                            Hora50N = RetornaDifData(INICIO, #11:59:00 PM#) + #12:01:00 AM# + FIM
                            End If
                    If tfim < #10:00:00 PM# And FIM <= #5:00:00 AM# Then
                    'se a jornada termina antes de 22:00 passa a computar depois das 22:00
                     'e como passou de 05:00 ele tem mais 05:00
                            If INICIO > #10:00:00 PM# Then
                                Hora50N = RetornaDifData(INICIO, #11:59:00 PM#) + #12:01:00 AM# + FIM
                                Else
                                Hora50N = #2:00:00 AM# + FIM
                            End If
                        End If
                    Else
                            If tfim >= #10:00:00 PM# And FIM > #5:00:00 AM# Then
                                'se a hora extra comeca depois de 22:00 ele tem 50% noturno
                                ' se ele passar de 05:00, ele também tem mais 05:00
                                    Hora50N = RetornaDifData(tfim, #11:59:00 PM#) + #5:01:00 AM#
                                    End If
                            If tfim < #10:00:00 PM# And FIM > #5:00:00 AM# Then
                            'se a jornada termina antes de 22:00 passa a computar depois das 22:00
                             'e como passou de 05:00 ele tem mais 05:00
                                        Hora50N = #7:00:00 AM#
                                     End If
                            If tfim >= #10:00:00 PM# And FIM <= #5:00:00 AM# Then
                                'se a hora extra comeca depois de 22:00 ele tem 50% noturno
                                ' se ele passar de 05:00, ele também tem mais 05:00
                                    Hora50N = RetornaDifData(tfim, #11:59:00 PM#) + #12:01:00 AM# + FIM
                                    End If
                            If tfim < #10:00:00 PM# And FIM <= #5:00:00 AM# Then
                            'se a jornada termina antes de 22:00 passa a computar depois das 22:00
                             'e como passou de 05:00 ele tem mais 05:00
                                        Hora50N = #2:00:00 AM# + FIM
                                    End If
                    End If
                
                
        Case 4 ' diurno e saiu diurno
        'falta comentar e concluir
            If FIM <= #10:00:00 PM# Then
                'funcionario saiu antes de 22:00 não tem hora noturna
                        Hora50N = #12:00:00 AM#
            Else
                'funcionário saiu depois de 22:00 pode receber hora noturna
                If FIM > tfim Then
                'funcionario saiu depois da hora e depois de 22:00
                        If tfim > #10:00:00 PM# Then
                        'saída do funcionario maior que 22:oo
                            If INICIO < tfim Then
                            'inicio menor que o termino do turno
                                Hora50N = RetornaDifData(tfim, FIM)
                                Else
                                'inicio maior ou igual a saída do turno
                                Hora50N = RetornaDifData(INICIO, FIM)
                                End If
                        Else
                                'turno de saída é menor do que 22:oo
                                If INICIO > #10:00:00 PM# And INICIO > tfim Then
                                    'inicio do funcionário é depois de 22:oo e depois de sua saída
                                            Hora50N = RetornaDifData(INICIO, FIM)
                                End If
                                If INICIO <= #10:00:00 PM# Then
                                    'inicio do funcionário é depois de 22:oo e depois de sua saída
                                            Hora50N = RetornaDifData(#10:00:00 PM#, FIM)
                                End If
                        End If
                End If
            End If
        End Select
End Select

'If Hora50N > Tot Then
'    MsgBox ("Retorno incorreto, consulte administrador")
'    End If

End Function

Function Hora100(FERIADO As Boolean, Proximodia As Boolean, INICIO As Date, FIM As Date, tinicio As Date, tfim As Date) As Date
Dim Tot As Date
Dim escolhaf As Integer
Dim escolhat As Integer

Tot = RetornaDifData(INICIO, FIM)
If FERIADO = True And Proximodia = True Then
    escolhaf = 1 'feriado e proximo dia feriado
    End If
If FERIADO = False And Proximodia = True Then
    escolhaf = 2 ' dia normal e proximo dia feriado
    End If
If FERIADO = True And Proximodia = False Then
    escolhaf = 3 ' feriado e proximo dia normal
    End If
If FERIADO = False And Proximodia = False Then
    escolhaf = 4 ' dia normal e proximo dia normal
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 1 ' funcionario noturno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 2 ' funcionario noturno e saiu diurno
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 3 ' funcionario é diurno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 4 ' funcionario é diurno e saiu diurno
    End If

Select Case escolhaf
Case 1 'feriado e proximo dia feriado
    Select Case escolhat
        Case 1  ' noturno saiu noturno
            If FIM > #5:00:00 AM# Then
                If INICIO < #10:00:00 PM# Then
                    Hora100 = RetornaDifData(INICIO, #10:00:00 PM#) + RetornaDifData(#5:00:00 AM#, FIM)
                    Else
                    Hora100 = RetornaDifData(#5:00:00 AM#, FIM)
                End If
            Else
            If INICIO < #10:00:00 PM# Then
                    Hora100 = RetornaDifData(INICIO, #10:00:00 PM#)
                End If
            End If
        Case 3  ' noturno saiu noturno
            If FIM > #5:00:00 AM# Then
                If INICIO < #10:00:00 PM# Then
                    Hora100 = RetornaDifData(INICIO, #10:00:00 PM#) + RetornaDifData(#5:00:00 AM#, FIM)
                    Else
                    Hora100 = RetornaDifData(#5:00:00 AM#, FIM)
                End If
            Else
            If INICIO < #10:00:00 PM# Then
                    Hora100 = RetornaDifData(INICIO, #10:00:00 PM#)
                End If
            End If
        
        Case 2  ' noturno e saiu diurno
            If Saida > #10:00:00 PM# Then
                If INICIO < #10:00:00 PM# Then
                    Hora100 = RetornaDifData(INICIO, #10:00:00 PM#)
                    End If
                Else
                    Hora100 = RetornaDifData(INICIO, FIM)
            End If
        Case 4  ' noturno e saiu diurno
            If Saida > #10:00:00 PM# Then
                If INICIO < #10:00:00 PM# Then
                    Hora100 = RetornaDifData(INICIO, #10:00:00 PM#)
                    End If
                Else
                    Hora100 = RetornaDifData(INICIO, FIM)
            End If
        End Select
Case 2 ' dia normal e proximo dia feriado
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            If FIM > tfim Then
                If tfim < #5:00:00 AM# Then
                    If FIM > #5:00:00 AM# Then
                        Hora100 = RetornaDifData(#5:00:00 AM#, FIM)
                    End If
                Else
                    If FIM > #5:00:00 AM# Then
                        Hora100 = RetornaDifData(tfim, FIM)
                    End If
                End If
            End If
        Case 2 ' noturno e saiu diurno
                Hora100 = #12:00:00 AM#
        Case 3 ' diurno e saiu noturno
            
                If FIM > #5:00:00 AM# Then
                'se a hora extra comeca depois de 00:00 ele não tem 100%n, mas pode receber
                       Hora100 = RetornaDifData(#5:00:00 AM#, FIM)
                    End If
                            
        Case 4 ' diurno e saiu diurno
                Hora100 = #12:00:00 AM#
        End Select

Case 3 ' feriado e proximo dia normal
    Select Case escolhat
        Case 1 ' noturno saiu noturno ou diurno e saiu noturno
           If INICIO < #10:00:00 PM# Then
                Hora100 = RetornaDifData(INICIO, #10:00:00 PM#)
                End If
        Case 3  ' noturno saiu noturno ou diurno e saiu noturno
           If INICIO < #10:00:00 PM# Then
                Hora100 = RetornaDifData(INICIO, #10:00:00 PM#)
                End If
        Case 2  ' noturno e saiu diurno
            If FIM > #10:00:00 PM# And INICIO < #10:00:00 PM# Then
                Hora100 = RetornaDifData(INICIO, #10:00:00 PM#)
                End If
            If FIM <= #10:00:00 PM# Then
                Hora100 = RetornaDifData(INICIO, FIM)
                End If
        Case 4 ' noturno e saiu diurno
            If FIM > #10:00:00 PM# And INICIO < #10:00:00 PM# Then
                Hora100 = RetornaDifData(INICIO, #10:00:00 PM#)
                End If
            If FIM <= #10:00:00 PM# Then
                Hora100 = RetornaDifData(INICIO, FIM)
                End If
        End Select

Case 4 ' dia normal e proximo dia normal

    Select Case escolhat
        Case 1 To 4
        Hora100 = #12:00:00 AM#
        End Select
        

End Select
If Hora100 > Tot Then
    MsgBox ("Retorno incorreto, consulte administrador")
    End If
End Function
Function Hora100N(FERIADO As Boolean, Proximodia As Boolean, INICIO As Date, FIM As Date, tinicio As Date, tfim As Date) As Date
Dim Tot As Date
Dim escolhaf As Integer
Dim escolhat As Integer

Tot = RetornaDifData(INICIO, FIM)
If FERIADO = True And Proximodia = True Then
    escolhaf = 1 'feriado e proximo dia feriado
    End If
If FERIADO = False And Proximodia = True Then
    escolhaf = 2 ' dia normal e proximo dia feriado
    End If
If FERIADO = True And Proximodia = False Then
    escolhaf = 3 ' feriado e proximo dia normal
    End If
If FERIADO = False And Proximodia = False Then
    escolhaf = 4 ' dia normal e proximo dia normal
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 1 ' funcionario noturno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = True Then
    escolhat = 2 ' funcionario noturno e saiu diurno
    End If
If RetornaDobra(INICIO, FIM) = True And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 3 ' funcionario é diurno e saiu noturno
    End If
If RetornaDobra(INICIO, FIM) = False And RetornaDobra(tinicio, tfim) = False Then
    escolhat = 4 ' funcionario é diurno e saiu diurno
    End If

Select Case escolhaf
Case 1 'feriado e proximo dia feriado
    Select Case escolhat
        Case 1  ' noturno saiu noturno
            If INICIO < #10:00:00 PM# Then
                If FIM > #5:00:00 AM# Then
                    Hora100N = #7:00:00 AM#
                    Else
                    Hora100N = #2:00:00 AM# + FIM
                End If
            Else
                If FIM > #5:00:00 AM# Then
                    Hora100N = #5:01:00 AM# + RetornaDifData(INICIO, #11:59:00 PM#)
                    Else
                    Hora100N = RetornaDifData(INICIO, #11:59:00 PM#) + FIM + #12:01:00 AM#
                End If
            End If
        
        Case 3 ' noturno saiu noturno
            If INICIO < #10:00:00 PM# Then
                If FIM > #5:00:00 AM# Then
                    Hora100N = #7:00:00 AM#
                    Else
                    Hora100N = #2:00:00 AM# + FIM
                End If
            Else
                If FIM > #5:00:00 AM# Then
                    Hora100N = #5:01:00 AM# + RetornaDifData(INICIO, #11:59:00 PM#)
                    Else
                    Hora100N = RetornaDifData(INICIO, #11:59:00 PM#) + FIM + #12:01:00 AM#
                End If
            End If
                        
        Case 2 ' noturno e saiu diurno
            If FIM > #10:00:00 PM# Then
                If INICIO < #10:00:00 PM# Then
                    Hora100N = RetornaDifData(#10:00:00 PM#, FIM)
                    Else
                    Hora100N = RetornaDifData(INICIO, FIM)
                End If
            End If
        Case 4  ' noturno e saiu diurno
            If FIM > #10:00:00 PM# Then
                If INICIO < #10:00:00 PM# Then
                    Hora100N = RetornaDifData(#10:00:00 PM#, FIM)
                    Else
                    Hora100N = RetornaDifData(INICIO, FIM)
                End If
            End If
        
        End Select
                
        
Case 2 ' dia normal e proximo dia feriado
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            If FIM > tfim Then
                If tfim < #5:00:00 AM# Then
                    If FIM > #5:00:00 AM# Then
                        Hora100N = RetornaDifData(tfim, #5:00:00 AM#)
                        Else
                        Hora100N = RetornaDifData(tfim, FIM)
                    End If
                End If
            End If
                    
                        
        Case 2 ' noturno e saiu diurno
                Hora100N = #12:00:00 AM#
        Case 3 ' diurno e saiu noturno
                If FIM >= #5:00:00 AM# Then
                'se a hora extra comeca depois de 22:00 ele não tem 50%, mas pode receber
                    Hora100N = #5:00:00 AM#
                    Else
                    Hora100N = FIM
                End If
                                
        Case 4 ' diurno e saiu diurno
                Hora100N = #12:00:00 AM#
        End Select
                
Case 3 'feriado e proximo dia normal
    Select Case escolhat
        Case 1 ' noturno saiu noturno
            If INICIO < #10:00:00 PM# Then
                Hora100N = #2:00:00 AM#
                Else
                Hora100N = RetornaDifData(INICIO, #11:59:00 PM#) + #12:01:00 AM#
                End If
        Case 3  ' noturno saiu noturno
            If INICIO < #10:00:00 PM# Then
                Hora100N = #2:00:00 AM#
                Else
                Hora100N = RetornaDifData(INICIO, #11:59:00 PM#) + #12:01:00 AM#
                End If
        Case 2 ' noturno e saiu diurno
            If FIM > #10:00:00 PM# Then
                Hora100N = RetornaDifData(#10:00:00 PM#, FIM)
                End If
        Case 4  ' noturno e saiu diurno
            If FIM > #10:00:00 PM# Then
                Hora100N = RetornaDifData(#10:00:00 PM#, FIM)
                End If
        
        End Select
        

Case 4 ' dia normal e proximo dia normal
  Select Case escolhat
        Case 1 To 4
        Hora100N = #12:00:00 AM#
        End Select

End Select

'If hora100n > Tot Then
'    MsgBox ("Retorno incorreto, consulte administrador")
'    End If

End Function
