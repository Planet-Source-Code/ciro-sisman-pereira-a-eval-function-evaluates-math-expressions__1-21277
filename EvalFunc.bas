Attribute VB_Name = "EvalFunc"
'
'  Avaliador de Expressões
'
'  (R) Ciro Sisman Pereira - 02/10/1999
'
'
'  Brasil - São Paulo
'

Option Explicit

Global Const gvValidTokens = "()0123456789+-/*.E"
Global Const NUMERO = 1
Global Const SINAL = 2

Type TokenType
 
     tipo  As Integer
     token As String

End Type

Global gvMsgError       As String
Global gvIsThereError   As Boolean
Global gvExpression     As String
Global gvTokenList(100) As TokenType
Global gvIdxToken       As Integer
Global gvEquacion       As String
Global gvPosI           As Integer
Global gvPosF           As Integer

Private Function retiraBrancos(aExpr As String) As String

  Dim iC      As Integer
  Dim newExpr As String

  aExpr = Trim(aExpr)
  newExpr = ""
  
  For iC = 1 To Len(aExpr)
  
      If Mid$(aExpr, iC, 1) <> " " Then
         newExpr = newExpr & Mid$(aExpr, iC, 1)
      End If
  
  Next iC

  retiraBrancos = newExpr

End Function

Public Function avaliarExpr(aExpr As String) As String
  
  Dim p1  As String
  Dim p2  As String
  
  gvIsThereError = False
  
  gvEquacion = UCase$(retiraBrancos(aExpr))
  If Len(gvEquacion) = 0 Then
     avaliarExpr = "*ERR-?"
     Exit Function
  End If
   
  Call existemTokensInvalidos
  If gvIsThereError Then GoTo OUT_ERROR
  
  Call verificaParenteses
  If gvIsThereError Then GoTo OUT_ERROR
  
  While True
  
       Call consegueExpressao
  
       If gvPosI = 1 Then
         Call preencheTokenList
         If gvIsThereError Then GoTo OUT_ERROR
         Call calcularExpressao
         If gvIsThereError Then GoTo OUT_ERROR
            gvEquacion = gvExpression + Right(gvEquacion, Len(gvEquacion) - gvPosF)
       ElseIf gvPosI = 0 Then
         If gvExpression = gvEquacion Then
            Call preencheTokenList
            If gvIsThereError Then GoTo OUT_ERROR
            Call calcularExpressao
            If gvIsThereError Then GoTo OUT_ERROR
            GoTo FINAL_PROC
         End If
       Else
          Call preencheTokenList
          If gvIsThereError Then GoTo OUT_ERROR
          Call calcularExpressao
          If gvIsThereError Then GoTo OUT_ERROR
          p1 = Left(gvEquacion, gvPosI - 1)
          p2 = Right(gvEquacion, Len(gvEquacion) - gvPosF)
          gvEquacion = p1 + gvExpression + p2
       End If
  
  Wend
  
FINAL_PROC:

  avaliarExpr = gvExpression
  
  Exit Function
  
OUT_ERROR:

  avaliarExpr = gvMsgError
  
End Function

Private Sub consegueExpressao()

  Dim iC      As Integer
  Dim sC      As String
  Dim oldExpr As String

  gvPosI = 0
  gvPosF = 0

  For iC = 1 To Len(gvEquacion)
  
      sC = Mid$(gvEquacion, iC, 1)
      
      If sC = "(" Then gvPosI = iC
      
      If sC = ")" Then
         gvPosF = iC
         GoTo MONTA_EXPR
      End If
      
  Next iC
  
MONTA_EXPR:
  
  If gvPosI > 0 Then
     gvExpression = Mid$(gvEquacion, gvPosI + 1, gvPosF - (gvPosI + 1))
  Else
     gvExpression = gvEquacion
  End If
  
End Sub

Private Sub verificaParenteses()

  Dim abertos  As Integer
  Dim fechados As Integer
  Dim iC       As Integer
  Dim sC       As String
  
  For iC = 1 To Len(gvEquacion)
     
      sC = Mid$(gvEquacion, iC, 1)
      
      If sC = "(" Then
         If abertos >= fechados Then
            abertos = abertos + 1
         Else
            gvIsThereError = True
            gvMsgError = "*ERR-()"
            Exit Sub
         End If
      End If
      
      If sC = ")" Then
         If abertos > fechados Then
            fechados = fechados + 1
         Else
            gvIsThereError = True
            gvMsgError = "*ERR-()"
            Exit Sub
         End If
      End If
      
  Next iC
  
  If abertos <> fechados Then
     gvIsThereError = True
     gvMsgError = "*ERR-()"
  End If
  
End Sub

Private Sub existemTokensInvalidos()

  Dim iC As Integer

  For iC = 1 To Len(gvEquacion)
      If InStr(1, gvValidTokens, Mid$(gvEquacion, iC, 1)) = 0 Then
         gvIsThereError = True
         gvMsgError = "*ERR-TOKEN"
         Exit Sub
      End If
  Next iC

End Sub

Private Sub preencheTokenList()

  Dim iC         As Integer
  Dim sC         As String
  Dim subTL(100) As TokenType
  Dim idxSub     As Integer

  gvIdxToken = 1
  idxSub = 1
  
  For iC = 1 To 100
      gvTokenList(iC).tipo = 0
      gvTokenList(iC).token = ""
      subTL(iC).tipo = 0
      subTL(iC).token = ""
  Next iC

  For iC = 1 To Len(gvExpression)
  
      sC = Mid$(gvExpression, iC, 1)
      
      Select Case sC
      
             Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "."
             
                  subTL(idxSub).tipo = NUMERO
                  subTL(idxSub).token = subTL(idxSub).token + sC
                  
             Case "+", "-", "/", "*", "E"
             
                  If subTL(idxSub).tipo > 0 Then idxSub = idxSub + 1
                  subTL(idxSub).tipo = SINAL
                  subTL(idxSub).token = sC
                  idxSub = idxSub + 1
                  
      End Select
  
  Next iC
  
  For iC = 1 To idxSub
      
     If iC = 1 And subTL(iC).tipo = SINAL Then
        If subTL(iC).token <> "-" Then
           gvIsThereError = True
           gvMsgError = "*ERR-SINAL"
           GoTo ERROR_FILL
        End If
     End If
     
     If subTL(iC).token = "-" Then
     
        If subTL(iC + 1).tipo = NUMERO Then
           gvTokenList(gvIdxToken).tipo = NUMERO
           gvTokenList(gvIdxToken).token = subTL(iC).token + subTL(iC + 1).token
           iC = iC + 1
           gvIdxToken = gvIdxToken + 1
        Else
           If subTL(iC + 1).token = "-" Then
              gvTokenList(gvIdxToken).tipo = NUMERO
              gvTokenList(gvIdxToken).token = subTL(iC + 2).token
              iC = iC + 2
              gvIdxToken = gvIdxToken + 1
           ElseIf subTL(iC + 1).token = "+" Then
              gvTokenList(gvIdxToken).tipo = NUMERO
              gvTokenList(gvIdxToken).token = "-" + subTL(iC + 2).token
              iC = iC + 2
              gvIdxToken = gvIdxToken + 1
           Else
              gvIsThereError = True
              gvMsgError = "*ERR-SINAL"
              GoTo ERROR_FILL
           End If
        End If
        
     ElseIf subTL(iC).token = "+" Then
     
        If subTL(iC + 1).tipo = NUMERO Then
           gvTokenList(gvIdxToken).tipo = NUMERO
           gvTokenList(gvIdxToken).token = subTL(iC).token + subTL(iC + 1).token
           iC = iC + 1
           gvIdxToken = gvIdxToken + 1
        Else
           If subTL(iC + 1).token = "-" Then
              gvTokenList(gvIdxToken).tipo = NUMERO
              gvTokenList(gvIdxToken).token = "-" + subTL(iC + 2).token
              iC = iC + 2
              gvIdxToken = gvIdxToken + 1
           ElseIf subTL(iC + 1).token = "+" Then
              gvTokenList(gvIdxToken).tipo = NUMERO
              gvTokenList(gvIdxToken).token = subTL(iC + 2).token
              iC = iC + 2
              gvIdxToken = gvIdxToken + 1
           Else
              gvIsThereError = True
              gvMsgError = "*ERR-SINAL"
              GoTo ERROR_FILL
           End If
        End If
     
     Else
         If subTL(iC).tipo = NUMERO Then
            gvTokenList(gvIdxToken).tipo = NUMERO
            gvTokenList(gvIdxToken).token = subTL(iC).token
            gvIdxToken = gvIdxToken + 1
         Else
            gvTokenList(gvIdxToken).tipo = SINAL
            gvTokenList(gvIdxToken).token = subTL(iC).token
            gvIdxToken = gvIdxToken + 1
         End If
         
     End If
     
  Next iC
  
  gvIdxToken = gvIdxToken - 1
  
ERROR_FILL:
  
End Sub

Private Sub calcularExpressao()

  Dim iC       As Integer
  Dim iD       As Integer
  Dim pos      As Integer
  Dim resp     As String
  Dim exp1     As String
  Dim exp2     As String
  Dim wrk      As Currency
    
  
  '*****************************************************
  '               Processa exponenciações
  '*****************************************************
  
  While True
  
       pos = posSinal("E", 1)
       If pos = 0 Then GoTo END_EXP
    
       exp1 = gvTokenList(pos - 1).token
       exp2 = gvTokenList(pos + 1).token
      
       resp = Trim(Str$(Val(exp1) ^ Val(exp2)))
       
       gvTokenList(pos - 1).token = Trim(resp)
       
       For iC = pos To pos + 1
           gvTokenList(iC).tipo = 0
           gvTokenList(iC).token = ""
       Next iC
            
       iD = pos
       For iC = pos + 2 To gvIdxToken
           gvTokenList(iD).tipo = gvTokenList(iC).tipo
           gvTokenList(iD).token = gvTokenList(iC).token
           gvTokenList(iC).tipo = 0
           gvTokenList(iC).token = ""
           iD = iD + 1
       Next iC
     
       gvIdxToken = gvIdxToken - 2
       
  Wend
  
END_EXP:
  
  '*****************************************************
  '               Processa divisões
  '*****************************************************
  
  While True
  
       pos = posSinal("/", 1)
       If pos = 0 Then GoTo END_DIV
    
       exp1 = gvTokenList(pos - 1).token
       exp2 = gvTokenList(pos + 1).token
      
       If Val(exp2) = 0 Then
          gvIsThereError = True
          gvMsgError = "*ERR-DIV"
          GoTo EXIT_BY_ERROR
       End If
      
       resp = Trim(Str$(Val(exp1) / Val(exp2)))
       
       gvTokenList(pos - 1).token = Trim(resp)
       
       For iC = pos To pos + 1
           gvTokenList(iC).tipo = 0
           gvTokenList(iC).token = ""
       Next iC
            
       iD = pos
       For iC = pos + 2 To gvIdxToken
           gvTokenList(iD).tipo = gvTokenList(iC).tipo
           gvTokenList(iD).token = gvTokenList(iC).token
           gvTokenList(iC).tipo = 0
           gvTokenList(iC).token = ""
           iD = iD + 1
       Next iC
     
       gvIdxToken = gvIdxToken - 2
       
  Wend
  
END_DIV:
  
  '*****************************************************
  '               Processa multiplicações
  '*****************************************************
  
  While True
  
       pos = posSinal("*", 1)
       If pos = 0 Then GoTo END_MUL
    
       exp1 = gvTokenList(pos - 1).token
       exp2 = gvTokenList(pos + 1).token
      
       resp = Trim(Str$(Val(exp1) * Val(exp2)))
       
       gvTokenList(pos - 1).token = Trim(resp)
       
       For iC = pos To pos + 1
           gvTokenList(iC).tipo = 0
           gvTokenList(iC).token = ""
       Next iC
            
       iD = pos
       For iC = pos + 2 To gvIdxToken
           gvTokenList(iD).tipo = gvTokenList(iC).tipo
           gvTokenList(iD).token = gvTokenList(iC).token
           gvTokenList(iC).tipo = 0
           gvTokenList(iC).token = ""
           iD = iD + 1
       Next iC
     
       gvIdxToken = gvIdxToken - 2
       
  Wend
  
END_MUL:
  
  '*****************************************************
  '               Processa somas/subtrações
  '*****************************************************
    
  If gvIdxToken > 1 Then
     For iC = 1 To gvIdxToken
         wrk = wrk + Val(gvTokenList(iC).token)
     Next iC
     gvTokenList(1).token = Trim(Str$(wrk))
  End If
    
  gvTokenList(1).token = Trim(Str$(CCur(Val(gvTokenList(1).token))))
    
  gvExpression = gvTokenList(1).token
  
EXIT_BY_ERROR:
    
End Sub

Private Function posSinal(aSinal As String, inic As Integer) As Integer

  Dim iC As Integer

  posSinal = 0

  For iC = inic To gvIdxToken
  
      If aSinal = gvTokenList(iC).token Then
         posSinal = iC
         GoTo END_POSSINAL
      End If
  
  Next iC
  
END_POSSINAL:

End Function
