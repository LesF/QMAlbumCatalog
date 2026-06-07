VERSION 5.00
Begin VB.Form FormServiceConfig 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "API Service Settings"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TextAccount 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox TextPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox TextUser 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox TextPort 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "4243"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox TextHost 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "QM Account"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hostname"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "FormServiceConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    TextHost.Text = SVC_Host
    TextPort.Text = SVC_Port
    TextUser.Text = SVC_User
    TextPass.Text = SVC_Pass
    TextAccount.Text = SVC_Account
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    '-----
    Dim userText As String
    Dim newPort As Long
    '-----
    userText = Trim(TextHost.Text)
    IniWrite "Connection", "Host", userText
    
    newPort = 4243
    userText = Trim(TextPort.Text)
    If IsNumeric(userText) Then newPort = CLng(userText)
    IniWrite "Connection", "Port", CStr(newPort)
    
    userText = Trim(TextUser.Text)
    IniWrite "Connection", "User", userText
    
    userText = Trim(TextPass.Text)
    IniWrite "Connection", "Password", userText
    
    userText = Trim(TextAccount.Text)
    IniWrite "Connection", "Account", userText

    Unload Me
End Sub

