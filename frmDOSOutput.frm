VERSION 5.00
Begin VB.Form frmDOSOutput 
   Caption         =   "DOS Outputs"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCommand 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   7095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   4080
      Width           =   1875
   End
   Begin VB.TextBox txtOutputs 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3135
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   720
      Width           =   7155
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   4080
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Command:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "frmDOSOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'DOSOutpus
'Capture the outputs of a DOS command
'Author: Marco Pipino
'marcopipino@libero.it
'28/02/2002


Option Explicit

Private WithEvents objDOS As DOSOutputs
Attribute objDOS.VB_VarHelpID = -1

Private Sub cmdExecute_Click()
    On Error GoTo errore
    objDOS.CommandLine = txtCommand.Text
    objDOS.ExecuteCommand
    Exit Sub
errore:
    MsgBox (Err.Description & " - " & Err.Source & " - " & CStr(Err.Number))
End Sub

Private Sub cmdExit_Click()
    Set objDOS = Nothing
    End
End Sub

Private Sub Form_Load()
    Set objDOS = New DOSOutputs
End Sub

Private Sub objDOS_ReceiveOutputs(CommandOutputs As String)
    txtOutputs.Text = txtOutputs.Text & CommandOutputs
End Sub

Private Sub txtOutputs_Change()
    txtOutputs.SelStart = Len(txtOutputs.Text)
End Sub
