VERSION 5.00
Begin VB.Form frmEvalTest 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Eval Test Window"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4395
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmExpr 
      Caption         =   "Result is"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   3735
      Begin VB.Label lblResposta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "&Process"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   3615
   End
   Begin VB.TextBox txtExpr 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblExpr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type an expression:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1725
   End
End
Attribute VB_Name = "frmEvalTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExec_Click()

  lblResposta.Caption = EvalFunc.avaliarExpr(txtExpr.Text)

  txtExpr.SetFocus

End Sub

Private Sub cmdSair_Click()

  End

End Sub
