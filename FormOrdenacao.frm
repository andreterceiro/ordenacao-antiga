VERSION 5.00
Begin VB.Form FormPrincipal 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordenação"
   ClientHeight    =   6390
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   7920
   Icon            =   "FormOrdenacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   7440
      Top             =   0
   End
   Begin VB.CommandButton cmdvoltar 
      BackColor       =   &H0000C000&
      Caption         =   "<<"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdavancar 
      BackColor       =   &H0000C000&
      Caption         =   ">>"
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdreiniciar 
      BackColor       =   &H0000C000&
      Caption         =   "Reiniciar"
      Height          =   495
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton CmdTerminei 
      BackColor       =   &H0000C000&
      Caption         =   "Terminei !"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton cmdlimpar 
      BackColor       =   &H0000C000&
      Caption         =   "Limpar"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5640
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Arrastar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   975
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      Height          =   2010
      Left            =   5040
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4080
      Width           =   2535
   End
   Begin VB.CommandButton Cmdcomparar 
      BackColor       =   &H0000C000&
      Caption         =   "Comparar"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Tempo"
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label lblsegundos 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "0"
      Height          =   255
      Left            =   3960
      TabIndex        =   29
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Respostas :"
      Height          =   255
      Left            =   600
      TabIndex        =   28
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lbltentativas 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "0"
      Height          =   255
      Left            =   1920
      TabIndex        =   27
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Comparações feitas"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Maior"
      Height          =   255
      Left            =   6840
      TabIndex        =   25
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Menor"
      Height          =   255
      Left            =   480
      TabIndex        =   24
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   23
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   22
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   21
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   20
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   495
      Index           =   4
      Left            =   3960
      TabIndex        =   19
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   495
      Index           =   5
      Left            =   4800
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   495
      Index           =   6
      Left            =   5640
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   495
      Index           =   7
      Left            =   6480
      TabIndex        =   16
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C000C0&
      Height          =   495
      Index           =   7
      Left            =   6480
      TabIndex        =   15
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Height          =   495
      Index           =   6
      Left            =   5640
      TabIndex        =   14
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Height          =   495
      Index           =   5
      Left            =   4800
      TabIndex        =   13
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   495
      Index           =   4
      Left            =   3960
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   3
      Left            =   3120
      TabIndex        =   11
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Height          =   495
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H008080FF&
      Height          =   495
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.Menu menger 
      Caption         =   "&Geral"
      Begin VB.Menu menrei 
         Caption         =   "Reiniciar"
         Shortcut        =   ^R
      End
      Begin VB.Menu mensai 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu menpon 
      Caption         =   "&Ponutauções"
      Begin VB.Menu menVisualizar 
         Caption         =   "Visualizar"
         Shortcut        =   ^P
      End
      Begin VB.Menu menLimpar 
         Caption         =   "Limpar"
      End
   End
   Begin VB.Menu menocu 
      Caption         =   "&Mostrar"
      Begin VB.Menu mencom 
         Caption         =   "Comparações"
         Checked         =   -1  'True
      End
      Begin VB.Menu mentem 
         Caption         =   "Tempo"
         Checked         =   -1  'True
      End
      Begin VB.Menu menlis 
         Caption         =   "Lista"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu menaju 
      Caption         =   "&Ajuda"
      Begin VB.Menu mentop 
         Caption         =   "Tópicos de ajuda"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mensob 
         Caption         =   "Sobre..."
      End
   End
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dado(1 To 12) As String
Dim ContadorApertados As Byte
Dim Vetor(7) As Byte
Dim VetorUsuario(7, 1) As Byte
Dim SPrimeiro As String
Dim SSegundo As String
Dim Primeiro As Byte
Dim Segundo As Byte

Private Sub Check1_Click()
Dim cont As Byte
For cont = 0 To 7
    Label1(cont).DragMode = Check1.Value
Next cont
    
End Sub

Private Sub cmdavancar_Click()
Dim cont As Integer
cont = 6
For cont = 6 To 0 Step -1
    If Label2(cont).BorderStyle = 1 And Label2(cont).BackColor <> &H8000000C Then
        Label2(cont + 1).BackColor = Label2(cont).BackColor
        Label2(cont).BackColor = &H8000000C
        VetorUsuario(cont + 1, 0) = VetorUsuario(cont, 0)
        VetorUsuario(cont, 0) = 255
        VetorUsuario(cont + 1, 1) = VetorUsuario(cont, 1)
    End If
        
Next cont
End Sub

Private Sub Cmdcomparar_Click()
Dim cont As Byte

If ContadorApertados <> 2 Then
    MsgBox "Favor escolha duas figuras para comparação", vbInformation, "Ops..."
Else
    Primeiro = 10
    For cont = 0 To 7
        If Label1(cont).BorderStyle = 1 Then
            If Primeiro = 10 Then
                Primeiro = cont
            Else
                Segundo = cont
                cont = 7
            End If
        End If
    Next cont
       
    ConverteValor
    
    If Vetor(Primeiro) = Vetor(Segundo) Then
       MsgBox SPrimeiro & " = " & SSegundo, vbInformation, "Resultado"
        List1.AddItem (SPrimeiro & " = " & SSegundo)
    ElseIf Vetor(Primeiro) < Vetor(Segundo) Then
       MsgBox SPrimeiro & " < " & SSegundo, vbInformation, "Resultado"
        List1.AddItem (SPrimeiro & " < " & SSegundo)
    Else
       MsgBox SPrimeiro & " > " & SSegundo, vbInformation, "Resultado"
        List1.AddItem (SPrimeiro & " > " & SSegundo)
    End If
lbltentativas.Caption = lbltentativas.Caption + 1
    If lbltentativas.Caption < CInt(dado(4)) Then
        lbltentativas.BackColor = 49152
    ElseIf lbltentativas.Caption < CInt(dado(6)) Then
        lbltentativas.BackColor = &H80FFFF
    Else
        lbltentativas.BackColor = &HFF&
    End If
End If
End Sub

Private Sub cmdlimpar_Click()
Dim cont As Byte
    For cont = 0 To 7
        If Label2(cont).BorderStyle = 1 Then Label2(cont).BackColor = &H8000000C
    Next cont
End Sub



Private Sub cmdreiniciar_Click()
    ReiniciarBotaoMenu
    Reiniciar

    
End Sub

Private Sub CmdTerminei_Click()
Dim cont As Byte
Dim acerto As Boolean
Dim repetido As Boolean
Dim recont As Byte
Dim branco As Boolean

repetido = False
acerto = True
branco = False

'-------------------------------
'    Teste
'-------------------------------
    'Dim tentativas As Integer
    'Dim tempo As Integer
'    lbltentativas.Caption = InputBox("Tentativas")
'    lblsegundos.Caption = InputBox("tempo")
'RegistrarRecorde ("teste")
'menpon_Click
    
'Exit Sub
'-------------------------------
'    Fim do teste
'-------------------------------

   cmdvoltar.Enabled = False
   cmdavancar.Enabled = False
    
    For cont = 0 To 6 Step 1
        If VetorUsuario(cont, 0) > VetorUsuario(cont + 1, 0) Then acerto = False
        If VetorUsuario(cont, 0) = 255 Then
            branco = True
        Else
            Label2(cont).Caption = VetorUsuario(cont, 0)
        End If
        Label1(cont).Caption = Vetor(cont)
                             
        If cont <> 0 Then
            For recont = 0 To cont - 1
                If VetorUsuario(cont, 1) = VetorUsuario(recont, 1) Then repetido = True
            Next recont
        End If
    
        Label1(cont).Enabled = False
        Label2(cont).Enabled = False
                
                
    Next cont

Label1(7).Enabled = False
Label2(7).Enabled = False
CmdTerminei.Enabled = False
cmdlimpar.Enabled = False
Check1.Enabled = False
Cmdcomparar.Enabled = False

Label1(7).Caption = Vetor(7)
    
If VetorUsuario(7, 0) <> 255 Then
    Label2(7).Caption = VetorUsuario(7, 0)
Else
    branco = True
End If

If branco = True Then
    Timer1.Enabled = False
    MsgBox "Você deixou pelo menos uma das células da resposta em branco !", vbCritical, "Branco"
    Exit Sub
End If
    
    For recont = 0 To 6
        If VetorUsuario(recont, 1) = VetorUsuario(7, 1) Then repetido = True
    Next recont

If repetido = True Then
    Timer1.Enabled = False
    MsgBox "Você repetiu pelo menos uma cor na resposta", vbCritical, "Erro"
    Exit Sub
End If



If Vetor(7) = 255 Then acerto = False
If acerto = True Then
    Timer1.Enabled = False
    MsgBox "Ordem certa !", vbInformation, "Parabéns !"
    If lblsegundos.BackColor <> &HFF Or lbltentativas.BackColor <> &HFF Then
        Dim strrecorde
        strrecorde = "xxxxxxxxxx"
        While (Not (Len(strrecorde) > 0 And Len(strrecorde) < 9)) And strrecorde <> ""
            strrecorde = InputBox("Você obteve um desempenho fantástico. Digite o nome que você deseja que apareça no placar (USE DE 1 A 8 CARACTERES)", "Parabéns !")
        Wend
    End If
    RegistrarRecorde (strrecorde)
    menVisualizar_Click
Else
    Timer1.Enabled = False
    MsgBox "Ordem errada !", vbCritical, "Erro !"
End If
    

End Sub

Private Sub cmdvoltar_Click()
Dim cont As Integer

For cont = 1 To 7 Step 1
    If Label2(cont).BorderStyle = 1 And Label2(cont).BackColor <> &H8000000C Then
        Label2(cont - 1).BackColor = Label2(cont).BackColor
        Label2(cont).BackColor = &H8000000C
        VetorUsuario(cont - 1, 0) = VetorUsuario(cont, 0)
        VetorUsuario(cont, 0) = 255
        VetorUsuario(cont - 1, 1) = VetorUsuario(cont, 1)
    End If
Next cont

End Sub

Private Sub Form_initialize()
   Randomize Timer
   Reiniciar
   IniciarRecordes
End Sub



Private Sub Label1_Click(Index As Integer)
If Label1(Index).BorderStyle = 0 Then
    If ContadorApertados < 2 Then
        Label1(Index).BorderStyle = 1
        ContadorApertados = ContadorApertados + 1
    End If
Else
    Label1(Index).BorderStyle = 0
    ContadorApertados = ContadorApertados - 1
End If
End Sub

Private Sub Label2_Click(Index As Integer)
If Label2(Index).BorderStyle = 0 Then
        Label2(Index).BorderStyle = 1
Else
    Label2(Index).BorderStyle = 0
End If
End Sub


Sub ConverteValor()
If Primeiro = 0 Then
    SPrimeiro = "Branco"
ElseIf Primeiro = 1 Then
    SPrimeiro = "Vermelho"
ElseIf Primeiro = 2 Then
    SPrimeiro = "Laranja"
ElseIf Primeiro = 3 Then
    SPrimeiro = "Amarelo"
ElseIf Primeiro = 4 Then
    SPrimeiro = "Verde"
ElseIf Primeiro = 5 Then
    SPrimeiro = "Azul claro"
ElseIf Primeiro = 6 Then
    SPrimeiro = "Azul escuro"
Else
    SPrimeiro = "Lilás"
End If

If Segundo = 0 Then
    SSegundo = "Branco"
ElseIf Segundo = 1 Then
    SSegundo = "Vermelho"
ElseIf Segundo = 2 Then
    SSegundo = "Laranja"
ElseIf Segundo = 3 Then
    SSegundo = "Amarelo"
ElseIf Segundo = 4 Then
    SSegundo = "Verde"
ElseIf Segundo = 5 Then
    SSegundo = "Azul claro"
ElseIf Segundo = 6 Then
    SSegundo = "Azul escuro"
Else
    SSegundo = "Lilás"
End If
End Sub



Private Sub Label2_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
If State = 0 Then
    If Source.Index = 0 Then
        Label2(Index).BackColor = &HFFFFFF
    ElseIf Source.Index = 1 Then
        Label2(Index).BackColor = &H8080FF
    ElseIf Source.Index = 2 Then
        Label2(Index).BackColor = &H80C0FF
    ElseIf Source.Index = 3 Then
        Label2(Index).BackColor = &HC0FFFF
    ElseIf Source.Index = 4 Then
        Label2(Index).BackColor = &HC0FFC0
    ElseIf Source.Index = 5 Then
        Label2(Index).BackColor = &HFFFFC0
    ElseIf Source.Index = 6 Then
        Label2(Index).BackColor = &HFF0000
    Else
       Label2(Index).BackColor = &HC000C0
    End If
End If
VetorUsuario(Index, 0) = Vetor(Source.Index)
VetorUsuario(Index, 1) = Source.Index

End Sub

Sub Reiniciar()

cmdvoltar.Enabled = True
cmdavancar.Enabled = True

Dim cont As Byte
    For cont = 0 To Val(Right(Time, 2)) + Val(Mid(Time, 4, 2))
    Rnd
    Next cont
    For cont = 0 To 7
        Vetor(cont) = Int(Rnd * (253) + 1)
        VetorUsuario(cont, 0) = 255
    Next cont
lblsegundos.Caption = 0
lblsegundos.BackColor = &HC000&
Timer1.Enabled = True
Check1.Value = 0
End Sub

Private Sub IniciarRecordes()
Dim caminho As String
Dim cont As Byte
caminho = App.Path & "\recordes.rec"
If Dir(caminho) <> "" Then
    Open caminho For Input As #1
    For cont = 1 To 12
    If EOF(1) Then
        cont = 12
    Else
        Input #1, dado(cont)
    End If
    Next cont
    Close #1
Else
   Open App.Path & "\recordes.rec" For Output As #2
   Print #2, "450" & vbCrLf & "550" & vbCrLf & "700" & vbCrLf & "20" & vbCrLf & "22" & vbCrLf & "25" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---"
   Close #2
   dado(1) = 450
   dado(2) = 550
   dado(3) = 700
   dado(4) = 20
   dado(5) = 22
   dado(6) = 25
End If
End Sub

Private Sub mencom_Click()
    mencom.Checked = Not (mencom.Checked)
    Label5.Visible = Not (Label5.Visible)
    lbltentativas.Visible = Not (lbltentativas.Visible)
End Sub

Private Sub menLimpar_Click()

 If MsgBox("Deseja realmente apagar as pontuações", vbDefaultButton2 + vbYesNo + vbQuestion, "Confirmação") = vbYes Then
   Open App.Path & "\recordes.rec" For Output As #2
      Print #2, "450" & vbCrLf & "550" & vbCrLf & "700" & vbCrLf & "20" & vbCrLf & "22" & vbCrLf & "25" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---" & vbCrLf & "---"
   Close #2
        
   dado(1) = 450
   dado(2) = 550
   dado(3) = 700
   dado(4) = 20
   dado(5) = 22
   dado(6) = 25
  
  Dim cont
   For cont = 7 To 12
      dado(cont) = "---"
   Next cont
 End If
 
End Sub

Private Sub menlis_Click()
    menlis.Checked = Not (menlis.Checked)
    List1.Visible = Not (List1.Visible)
End Sub

Private Sub menrei_Click()
    ReiniciarBotaoMenu
    Reiniciar
End Sub

Private Sub mensai_Click()
    End
End Sub

Private Sub mensob_Click()
    FormSobre.Show
End Sub

Private Sub mentem_Click()
    mentem.Checked = Not (mentem.Checked)
    Label7.Visible = Not (Label7.Visible)
    lblsegundos.Visible = Not (lblsegundos.Visible)
End Sub

Private Sub mentop_Click()
'    Shell "c:\arquivos de programas\internet explorer\iexplore.exe D:\backup\backup\486\Projetos\ordenacao\ajuda\help.htm"
 Dim strArquivo As String
 Dim objAjuda As ClasseAjuda
 
    Set objAjuda = New ClasseAjuda
    
    strArquivo = App.Path & "\ajuda\ajuda.chm"
    Call objAjuda.Show(strArquivo) ', "janelaHelp")
    Set objAjuda = Nothing

'   formAjuda.Show 1
End Sub

Private Sub menVisualizar_Click()
    MsgBox "As melhoeres pontuações são: " & vbCrLf & vbCrLf & "   Tempo:" & vbCrLf & "1º " & dado(7) & " - " & dado(1) & vbCrLf & "2º " & dado(8) & " - " & dado(2) & vbCrLf & "3º " & dado(9) & " - " & dado(3) & vbCrLf & vbCrLf & "   Comparações:" & vbCrLf & "1º " & dado(10) & " - " & dado(4) & vbCrLf & "2º " & dado(11) & " - " & dado(5) & vbCrLf & "3º " & dado(12) & " - " & dado(6), vbInformation + vbOKOnly, "Recordes"
End Sub

Private Sub Timer1_Timer()
    lblsegundos.Caption = lblsegundos.Caption + 1
    If lblsegundos.Caption < CDbl(dado(1)) Then
        lblsegundos.BackColor = &HC000&
    ElseIf lblsegundos.Caption < CDbl(dado(3)) Then
        lblsegundos.BackColor = &H80FFFF
    Else
        lblsegundos.BackColor = &HFF&
    End If
End Sub

Sub RegistrarRecorde(strecorde As String)
    If lblsegundos.Caption < CDbl(dado(2)) Then
        If lblsegundos.Caption < CDbl(dado(1)) Then
            dado(3) = dado(2)
            dado(9) = dado(8)
            dado(2) = dado(1)
            dado(8) = dado(7)
            dado(1) = lblsegundos.Caption
            dado(7) = strecorde
        Else
            dado(3) = dado(2)
            dado(9) = dado(8)
            dado(2) = lblsegundos.Caption
            dado(8) = strecorde
        End If
    Else
        If lblsegundos.Caption < CDbl(dado(3)) Then
            dado(3) = lblsegundos.Caption
            dado(9) = strecorde
        End If
    End If
    
    If lbltentativas.Caption < CDbl(dado(5)) Then
        If lbltentativas.Caption < CDbl(dado(4)) Then
            dado(6) = dado(5)
            dado(12) = dado(11)
            dado(5) = dado(4)
            dado(11) = dado(10)
            dado(4) = lbltentativas.Caption
            dado(10) = strecorde
        Else
            dado(6) = dado(5)
            dado(12) = dado(11)
            dado(5) = lbltentativas.Caption
            dado(11) = strecorde
        End If
    Else
        If lbltentativas.Caption < CDbl(dado(6)) Then
            dado(6) = lbltentativas.Caption
            dado(12) = strecorde
        End If
    End If

Open App.Path & "\recordes.rec" For Output As #3
Dim cont As Byte
    For cont = 1 To 12
        Print #3, dado(cont)
    Next cont
Close #3
End Sub


Sub ReiniciarBotaoMenu()
Dim cont As Byte
    For cont = 0 To 7
        Label2(cont).Caption = ""
        Label1(cont).Caption = ""
        Label1(cont).Enabled = True
        Label2(cont).Enabled = True
        Label1(cont).BorderStyle = 0
        Label2(cont).BorderStyle = 0
        Label2(cont).BackColor = &H8000000C
        lbltentativas.Caption = "0"
        If dado(4) > 0 Then
            lbltentativas.BackColor = &HC000&
        ElseIf dado(6) > 0 Then
            lbltentativas.BackColor = &H80FFFF
        Else
            lbltentativas.BackColor = &HFF&
        End If
    Next cont
    List1.Clear
    CmdTerminei.Enabled = True
    cmdlimpar.Enabled = True
    Check1.Enabled = True
    Cmdcomparar.Enabled = True
    ContadorApertados = 0
End Sub
