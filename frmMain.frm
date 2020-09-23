VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cd Player By: mikecanejo@hotmail.com"
   ClientHeight    =   2340
   ClientLeft      =   5100
   ClientTop       =   1425
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2340
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Track Jump"
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3255
      Begin VB.CommandButton Command1 
         Caption         =   "Play"
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Stop"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   720
         TabIndex        =   3
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Track:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open Tray"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Tray"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame frmTime 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Time Jump"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command7 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   "5"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seconds"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   600
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim Snd As CDAudio

Private Sub Command1_Click()
    Snd.SeekCDtoX Val(Text1)
End Sub

Private Sub Command2_Click()
    s$ = Snd.GetCDLength
    MsgBox "Total length of CD: " & s$, , "CD len"
End Sub

Private Sub Command3_Click()
    Snd.CloseCD
End Sub

Private Sub Command4_Click()
    Snd.EjectCD
End Sub

Private Sub Command5_Click()
    Snd.StopPlay
End Sub

Private Sub Command6_Click()
    Snd.ReWind Val(Text2) * 1000
End Sub

Private Sub Command7_Click()
    Snd.FastForward Val(Text2) * 1000
End Sub

Private Sub Form_Load()
    Set Snd = New CDAudio
    Snd.ReadyDevice
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Snd.StopPlay
    Snd.UnloadAll
End Sub

