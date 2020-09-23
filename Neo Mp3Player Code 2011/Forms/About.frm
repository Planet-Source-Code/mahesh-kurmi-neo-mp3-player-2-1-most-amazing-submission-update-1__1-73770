VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   Caption         =   " About Mahesh Mp3 Player ..."
   ClientHeight    =   4050
   ClientLeft      =   3345
   ClientTop       =   2370
   ClientWidth     =   6075
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   5250
      Top             =   1365
   End
   Begin VB.PictureBox picScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1575
      Left            =   0
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   396
      TabIndex        =   0
      Top             =   240
      Width           =   5940
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "http://neoaudioplayer.blogspot.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Left            =   2160
         MouseIcon       =   "About.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1200
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ScrollText As String
' scrolltext is the text to br scrolled
Dim rt As Long
Dim DrawingRect As RECT
'RECT is rectangle coordinates structure(type)
Dim UpperX As Long, UpperY As Long    'Upper left point PICSCROLL
Attribute UpperY.VB_VarUserMemId = 1073938435
Dim RectHeight As Long
Attribute RectHeight.VB_VarUserMemId = 1073938437

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub RunMain()

    Const IntervalTime As Long = 60
    frmAbout.Refresh
    rt = DrawText(picScroll.hDC, ScrollText, -1, DrawingRect, DT_CALCRECT)

    If rt = 0 Then
    Else
        'Drawing rect is rect TYpe variable
        DrawingRect.Top = picScroll.ScaleHeight
        DrawingRect.Left = 0
        DrawingRect.Right = picScroll.ScaleWidth
        RectHeight = DrawingRect.Bottom
        DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
    End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Load()
    On Error Resume Next

    bolAcercaShow = True

    ' Me.Caption = LineLanguage(40)
    ' Me.Icon = frmMain.Icon
    ScrollText = "........................................." & vbCrLf & _
                 "Mahesh Mp3 Player " & vbCrLf & "........................................." & vbCrLf & "VERSION 2.0" & vbCrLf & vbCrLf & _
                 "DEVELOPED BY:" & vbCrLf & _
                 "<< Mahesh Kurmi >>" & vbCrLf & _
                 "55 Shantinagar Indore" & vbCrLf & _
                 "M.P. INDIA" & vbCrLf & vbCrLf & _
                 "SEPTEMBER 2006" & vbCrLf & vbCrLf & _
                 "If you have any ideas," & vbCrLf & _
                 "comments, doubts, suggestions," & vbCrLf & _
                 "bugs, skins, languages, etc," & vbCrLf & _
                 "please email me." & vbCrLf & vbCrLf & _
                 "E-mail :" & vbCrLf & _
                 "mahesh_kurmi2003@yahoo.com" & vbCrLf & vbCrLf & _
                 "Web Site :" & vbCrLf & _
                 "www.Shikharclasses.com" & vbCrLf & _
                 "........................................." & vbCrLf & _
                 ".................CREDITS................." & vbCrLf & _
                 "........................................." & vbCrLf & _
                 "Idea and Start: Mortinez Corp: " & vbCrLf & _
                 "Big Thanks to LaVolpe and Arne Ester " & vbCrLf & _
                 "Some Controls: Thanks to VkControls" & vbCrLf & _
                 "THANKS PSC" & vbCrLf & _
                 "........................................." & vbCrLf & _
                 "........................................."


    Timer1.Enabled = True
    Me.Left = (Screen.Width - Me.Width) / 2   '// centrar formulario
    Me.Top = (Screen.Height - Me.Height) / 2
    picScroll.Move 3, 3, Me.ScaleWidth, Me.ScaleHeight

    RunMain    '// empezar a mover el texto
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFF00&
    Label1.FontUnderline = False

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picScroll.Width = Me.ScaleWidth - 2
    picScroll.Height = Me.ScaleHeight - 2
    Label1.Top = Me.ScaleHeight - 40
    Label1.Left = Me.ScaleWidth / 2 - Label1.Width / 2
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    bolAcercaShow = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFF00&
    Label1.FontUnderline = False

End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'for 3 d effect of pressing the label
    Label1.Move Label1.Left + 1, Label1.Top + 1
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFFFFC0
    Label1.FontUnderline = True
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim lngRETURN As Long
    Label1.ForeColor = &HFF00&
    Label1.FontUnderline = False
    Label1.Move Label1.Left - 1, Label1.Top - 1
    lngRETURN = ShellExecute(Me.hwnd, "Open", Label1.Caption, "", "", vbNormalFocus)
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub picScroll_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub picScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.ForeColor = &HFF00&
    Label1.FontUnderline = False

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Timer1_Timer()
    picScroll.cls
    DrawText picScroll.hDC, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
    DrawingRect.Top = DrawingRect.Top - 1
    DrawingRect.Bottom = DrawingRect.Bottom - 1
    If DrawingRect.Top < -(RectHeight) Then    '// Tiempo de reinicio
        DrawingRect.Top = picScroll.ScaleHeight
        DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
    End If

    picScroll.Refresh
    DoEvents
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
