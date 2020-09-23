VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configure Particle - Event Horizon"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3480
      Top             =   3720
   End
   Begin VB.CheckBox Check1 
      Caption         =   "More Parameters"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters:"
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox Check2 
         Caption         =   " Hold"
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2415
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CheckBox Check3 
            Caption         =   "Hold"
            Height          =   255
            Left            =   2400
            TabIndex        =   18
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   375
            Left            =   960
            TabIndex        =   16
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   375
            Left            =   960
            TabIndex        =   14
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   375
            Left            =   960
            TabIndex        =   12
            Top             =   960
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   960
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   960
            TabIndex        =   10
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Factor:"
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Phase:"
            Height          =   375
            Left            =   0
            TabIndex        =   13
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Angle:"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "NewY:"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "NewX:"
            Height          =   195
            Left            =   0
            TabIndex        =   7
            Top             =   120
            Width           =   480
         End
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Step:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label2 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   2775
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check1_Click()
Timer1.Enabled = CBool(Check1.Value)
Frame2.Visible = CBool(Check1.Value)
End Sub

Private Sub Check2_Click()
holdStep = CBool(Check2.Value)
End Sub

Private Sub Check3_Click()
holdFactor = CBool(Check2.Value)
SaveSetting App.EXEName, "Settings", "holdFactor", holdFactor

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Text1.Text = CStr(Step)
Label2.Caption = "Particle.dll, version " & App.Major & "." & App.Minor & ", build " & App.Revision
End Sub

Private Sub Text1_Change()
On Error Resume Next
Step = CSng(Text1.Text)
End Sub

Private Sub Text4_Change()
On Error Resume Next
angle = CSng(Text4.Text)

End Sub

Private Sub Text6_Change()
On Error Resume Next
Factor = CSng(Text6.Text)

End Sub

Private Sub Timer1_Timer()
Text1.Text = Step
Text2.Text = NewX
Text3.Text = NewY
Text4.Text = angle
Text5.Text = phase
Text6.Text = Factor
End Sub
