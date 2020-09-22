VERSION 5.00
Begin VB.Form frmRGCC 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraOperator 
      Height          =   1755
      Left            =   120
      TabIndex        =   6
      Top             =   420
      Width           =   1935
      Begin VB.OptionButton optOperator 
         Caption         =   "Not (Val A Only)"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   1380
         Width           =   1455
      End
      Begin VB.OptionButton optOperator 
         Caption         =   "Eqv"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1140
         Width           =   1215
      End
      Begin VB.OptionButton optOperator 
         Caption         =   "Imp"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   900
         Width           =   1215
      End
      Begin VB.OptionButton optOperator 
         Caption         =   "And"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   1215
      End
      Begin VB.OptionButton optOperator 
         Caption         =   "Or"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   420
         Width           =   1215
      End
      Begin VB.OptionButton optOperator 
         Caption         =   "XOr"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   1215
      End
   End
   Begin VB.TextBox txtOut 
      Height          =   315
      Left            =   900
      TabIndex        =   4
      Text            =   "0"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtB 
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      Text            =   "0"
      Top             =   60
      Width           =   1215
   End
   Begin VB.TextBox txtA 
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Text            =   "0"
      Top             =   60
      Width           =   1215
   End
   Begin VB.Label lblOut 
      Caption         =   "Out Value:"
      Height          =   255
      Left            =   60
      TabIndex        =   5
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblValB 
      Caption         =   "Value B:"
      Height          =   255
      Left            =   2220
      TabIndex        =   3
      Top             =   60
      Width           =   795
   End
   Begin VB.Label lblValA 
      Caption         =   "Value A:"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "frmRGCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub optOperator_Click(Index As Integer)
  If (txtA < 256) And (txtA > -1) And (txtB < 256) And (txtB > -1) Then
    txtA.Enabled = False
    txtB.Enabled = False
    fraOperator.Enabled = False
    txtOut.Enabled = False
    Select Case Index
      Case 0: txtOut = Logic_Oper(CByte(txtA), CByte(txtB), Operator_XOr)
      Case 1: txtOut = Logic_Oper(CByte(txtA), CByte(txtB), Operator_OR)
      Case 2: txtOut = Logic_Oper(CByte(txtA), CByte(txtB), Operator_And)
      Case 3: txtOut = Logic_Oper(CByte(txtA), CByte(txtB), Operator_Imp)
      Case 4: txtOut = Logic_Oper(CByte(txtA), CByte(txtB), Operator_Eqv)
      Case 5: txtOut = Logic_Oper(CByte(txtA), 0, Operator_Not)
    End Select
    txtA.Enabled = True
    txtB.Enabled = True
    fraOperator.Enabled = True
    txtOut.Enabled = True
  Else
    MsgBox "Input values must be between 0 and 255"
  End If
End Sub
