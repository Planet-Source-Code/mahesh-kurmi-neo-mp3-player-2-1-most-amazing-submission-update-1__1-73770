VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MaheshMp3 Player"
   ClientHeight    =   6945
   ClientLeft      =   3420
   ClientTop       =   3105
   ClientWidth     =   10320
   ForeColor       =   &H00000000&
   HasDC           =   0   'False
   Icon            =   "MMp3Player.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   10320
   Visible         =   0   'False
   Begin VB.Frame FrameContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9885
      Begin VB.PictureBox PicConfigInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   78
         Top             =   120
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.PictureBox picPlayInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4680
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   77
         Top             =   120
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.ListBox ListEq 
         Appearance      =   0  'Flat
         Height          =   810
         ItemData        =   "MMp3Player.frx":000C
         Left            =   4560
         List            =   "MMp3Player.frx":000E
         TabIndex        =   75
         Top             =   3840
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.PictureBox PicMIXER 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4800
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   153
         TabIndex        =   70
         Top             =   600
         Visible         =   0   'False
         Width           =   2295
         Begin MMPlayerXProject.Button Button 
            Height          =   105
            Index           =   19
            Left            =   120
            TabIndex        =   71
            ToolTipText     =   "Echo/Rec EQ"
            Top             =   120
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
            ButtonColor     =   -2147483643
            Selected        =   -1  'True
            Style           =   1
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   105
            Index           =   20
            Left            =   120
            TabIndex        =   72
            ToolTipText     =   "Echo/Rec EQ"
            Top             =   240
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
            ButtonColor     =   -2147483643
            Style           =   1
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   105
            Index           =   21
            Left            =   1830
            TabIndex        =   73
            ToolTipText     =   "Output File settings"
            Top             =   120
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
            ButtonColor     =   -2147483643
            Style           =   1
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   105
            Index           =   22
            Left            =   1830
            TabIndex        =   74
            Top             =   240
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   185
            ButtonColor     =   -2147483643
            Style           =   1
         End
      End
      Begin VB.Timer Timer_Intro 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1800
         Top             =   4200
      End
      Begin VB.Timer Timer_Player 
         Enabled         =   0   'False
         Interval        =   35
         Left            =   2280
         Top             =   4200
      End
      Begin VB.Timer Timer_Crossfade 
         Enabled         =   0   'False
         Interval        =   400
         Left            =   3240
         Top             =   4200
      End
      Begin VB.Timer Timer_Wait 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   2760
         Top             =   4230
      End
      Begin VB.PictureBox EqBackpic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   1755
         Left            =   4680
         ScaleHeight     =   117
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   170
         TabIndex        =   53
         Top             =   1140
         Width           =   2550
         Begin VB.PictureBox PicEqAmp 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            FillColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   60
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   163
            TabIndex        =   54
            Top             =   1380
            Width           =   2445
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   0
            Left            =   480
            TabIndex        =   55
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Vol_Eq_SliderCtrl 
            Height          =   915
            Left            =   60
            TabIndex        =   56
            Top             =   420
            Visible         =   0   'False
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   195
            Index           =   16
            Left            =   660
            TabIndex        =   57
            ToolTipText     =   "EQ On/Off"
            Top             =   60
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   344
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   195
            Index           =   17
            Left            =   60
            TabIndex        =   58
            ToolTipText     =   "Eq Free Movement"
            Top             =   60
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   344
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   195
            Index           =   18
            Left            =   360
            TabIndex        =   59
            ToolTipText     =   "Open DFX Editor"
            Top             =   60
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   344
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Eq_SliderCtrl balance 
            Height          =   105
            Left            =   1260
            TabIndex        =   60
            ToolTipText     =   "Balance"
            Top             =   150
            Width           =   1215
            _ExtentX        =   185
            _ExtentY        =   2143
            Max             =   255
            Value           =   128
            Position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   1
            Left            =   690
            TabIndex        =   61
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   2
            Left            =   900
            TabIndex        =   62
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   3
            Left            =   1110
            TabIndex        =   63
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   4
            Left            =   1320
            TabIndex        =   64
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   5
            Left            =   1530
            TabIndex        =   65
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   6
            Left            =   1740
            TabIndex        =   66
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   7
            Left            =   1950
            TabIndex        =   67
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   8
            Left            =   2160
            TabIndex        =   68
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
         Begin MMPlayerXProject.Eq_SliderCtrl Eq_SliderCtrl 
            Height          =   915
            Index           =   9
            Left            =   2370
            TabIndex        =   69
            Top             =   420
            Width           =   120
            _ExtentX        =   1614
            _ExtentY        =   212
            Min             =   -10
            Max             =   10
         End
      End
      Begin VB.PictureBox picWithoutEq 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   0
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   307
         TabIndex        =   49
         Top             =   1800
         Visible         =   0   'False
         Width           =   4635
      End
      Begin VB.PictureBox picNormalMode 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1740
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   116
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   311
         TabIndex        =   13
         Top             =   0
         Width           =   4665
         Begin VB.PictureBox PicInfo 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   165
            Left            =   120
            ScaleHeight     =   11
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   13
            TabIndex        =   83
            Top             =   480
            Width           =   195
         End
         Begin VB.PictureBox PicaTop 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   3720
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   82
            ToolTipText     =   "Always on Top"
            Top             =   870
            Width           =   375
         End
         Begin VB.PictureBox PicShuffle 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   75
            Left            =   4200
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   81
            ToolTipText     =   "Shuffle On/Off"
            Top             =   600
            Width           =   375
         End
         Begin VB.PictureBox PicRepeat 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   75
            Left            =   4200
            ScaleHeight     =   5
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   80
            ToolTipText     =   "Repeat On/Off"
            Top             =   720
            Width           =   375
         End
         Begin VB.PictureBox PicCrossfade 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   4200
            ScaleHeight     =   9
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   79
            ToolTipText     =   "Crossfade On/Off"
            Top             =   870
            Width           =   375
         End
         Begin VB.PictureBox picTemp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   3120
            ScaleHeight     =   27
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   37
            TabIndex        =   76
            Top             =   600
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.PictureBox picSpectrum 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   30
            ScaleHeight     =   28
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   133
            TabIndex        =   16
            Top             =   420
            Width           =   1995
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   14
            ToolTipText     =   "Previous track"
            Top             =   1230
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.ScrollText ScrollText 
            Height          =   90
            Index           =   0
            Left            =   30
            TabIndex        =   15
            Top             =   270
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   159
            BackColor       =   16777215
            CaptionText     =   " "
            AlignText       =   1
            ScrollVelocity  =   150
         End
         Begin MMPlayerXProject.ScrollText ScrollText 
            Height          =   90
            Index           =   2
            Left            =   960
            TabIndex        =   17
            Top             =   0
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   159
            BackColor       =   16777215
            CaptionText     =   " "
            AlignText       =   1
            ScrollType      =   1
            ScrollVelocity  =   150
            Scroll          =   -1  'True
         End
         Begin MMPlayerXProject.ScrollText ScrollText 
            Height          =   90
            Index           =   3
            Left            =   570
            TabIndex        =   18
            Top             =   270
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   159
            BackColor       =   16777215
            CaptionText     =   " "
            AlignText       =   1
            ScrollType      =   1
            ScrollVelocity  =   150
         End
         Begin MMPlayerXProject.ScrollText ScrollText 
            Height          =   90
            Index           =   1
            Left            =   1560
            TabIndex        =   19
            Top             =   270
            Width           =   3045
            _ExtentX        =   3043
            _ExtentY        =   159
            BackColor       =   16777215
            CaptionText     =   " "
            AlignText       =   2
            ScrollType      =   1
            ScrollVelocity  =   150
            Scroll          =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   300
            Index           =   1
            Left            =   480
            TabIndex        =   20
            ToolTipText     =   "Play"
            Top             =   1230
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   300
            Index           =   2
            Left            =   900
            TabIndex        =   21
            ToolTipText     =   "pause"
            Top             =   1230
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   300
            Index           =   3
            Left            =   1320
            TabIndex        =   22
            ToolTipText     =   "Stop"
            Top             =   1230
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   300
            Index           =   4
            Left            =   1740
            TabIndex        =   23
            ToolTipText     =   "Next"
            Top             =   1230
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   135
            Index           =   5
            Left            =   2700
            TabIndex        =   24
            ToolTipText     =   "Intro"
            Top             =   690
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   238
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   135
            Index           =   6
            Left            =   2340
            TabIndex        =   25
            ToolTipText     =   "Mute"
            Top             =   690
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   238
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   135
            Index           =   7
            Left            =   2880
            TabIndex        =   26
            ToolTipText     =   "Repeat"
            Top             =   690
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   238
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   135
            Index           =   8
            Left            =   2520
            TabIndex        =   27
            ToolTipText     =   "Randomize"
            Top             =   690
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   238
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   9
            Left            =   3420
            TabIndex        =   28
            ToolTipText     =   "Peset<<"
            Top             =   1380
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   10
            Left            =   3000
            TabIndex        =   29
            ToolTipText     =   "DFX"
            Top             =   1380
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   11
            Left            =   3840
            TabIndex        =   30
            ToolTipText     =   ">>Preset"
            Top             =   1380
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   12
            Left            =   15
            TabIndex        =   31
            ToolTipText     =   "Menu"
            Top             =   0
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   13
            Left            =   3960
            TabIndex        =   32
            ToolTipText     =   "Minimize"
            Top             =   30
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   14
            Left            =   4170
            TabIndex        =   33
            ToolTipText     =   "change Mode"
            Top             =   30
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   15
            Left            =   4380
            TabIndex        =   34
            ToolTipText     =   "Close"
            Top             =   30
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   135
            Index           =   23
            Left            =   2160
            TabIndex        =   35
            ToolTipText     =   "Configuration"
            Top             =   690
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   238
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Eq_SliderCtrl slider 
            Height          =   135
            Index           =   1
            Left            =   2160
            TabIndex        =   36
            ToolTipText     =   "Volume"
            Top             =   420
            Width           =   1155
            _ExtentX        =   238
            _ExtentY        =   2037
            Max             =   255
            Value           =   50
            Position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl slider 
            Height          =   135
            Index           =   0
            Left            =   30
            TabIndex        =   37
            ToolTipText     =   "Seek Bar"
            Top             =   1050
            Width           =   4575
            _ExtentX        =   238
            _ExtentY        =   8070
            Max             =   0
            Position        =   1
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   24
            Left            =   4260
            TabIndex        =   45
            ToolTipText     =   "EQ show/hide"
            Top             =   1380
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   265
            ButtonColor     =   65535
            Style           =   1
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   300
            Index           =   25
            Left            =   2160
            TabIndex        =   46
            ToolTipText     =   "Add tracks"
            Top             =   1230
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   529
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button Button 
            Height          =   150
            Index           =   26
            Left            =   2580
            TabIndex        =   47
            ToolTipText     =   "Playlist show/hide"
            Top             =   1380
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Eq_SliderCtrl slider 
            Height          =   135
            Index           =   5
            Left            =   3810
            TabIndex        =   50
            ToolTipText     =   "Balance"
            Top             =   420
            Width           =   795
            _ExtentX        =   238
            _ExtentY        =   1402
            Max             =   255
            Value           =   128
            Position        =   1
         End
         Begin VB.PictureBox PicActivate 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   -240
            ScaleHeight     =   14
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   323
            TabIndex        =   48
            Top             =   0
            Width           =   4845
         End
      End
      Begin VB.PictureBox picMiniMode 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         FillColor       =   &H00642909&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   0
         MousePointer    =   99  'Custom
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   483
         TabIndex        =   1
         Top             =   3060
         Visible         =   0   'False
         Width           =   7245
         Begin VB.PictureBox PicSpectrumMini 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   3060
            ScaleHeight     =   25
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   93
            TabIndex        =   51
            Top             =   120
            Width           =   1395
         End
         Begin MMPlayerXProject.ScrollText ScrollText 
            Height          =   90
            Index           =   4
            Left            =   4590
            TabIndex        =   2
            Top             =   390
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   159
            CaptionText     =   " "
            AlignText       =   1
            ScrollVelocity  =   150
         End
         Begin MMPlayerXProject.ScrollText ScrollText 
            Height          =   90
            Index           =   5
            Left            =   4590
            TabIndex        =   3
            Top             =   120
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   159
            BackColor       =   16777215
            CaptionText     =   " "
            AlignText       =   2
            ScrollType      =   1
            ScrollVelocity  =   150
            Scroll          =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   255
            Index           =   0
            Left            =   30
            TabIndex        =   4
            Top             =   240
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   450
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   450
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   255
            Index           =   2
            Left            =   690
            TabIndex        =   6
            ToolTipText     =   "Pause"
            Top             =   240
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   450
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   255
            Index           =   3
            Left            =   1020
            TabIndex        =   7
            ToolTipText     =   "Stop"
            Top             =   240
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   450
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   255
            Index           =   4
            Left            =   1350
            TabIndex        =   8
            ToolTipText     =   "next Track"
            Top             =   240
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   450
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   150
            Index           =   5
            Left            =   30
            TabIndex        =   9
            ToolTipText     =   "Open"
            Top             =   30
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   105
            Index           =   6
            Left            =   6780
            TabIndex        =   10
            ToolTipText     =   "Mute"
            Top             =   105
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   185
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   105
            Index           =   7
            Left            =   6930
            TabIndex        =   11
            ToolTipText     =   "Repeat"
            Top             =   105
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   185
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   105
            Index           =   8
            Left            =   7080
            TabIndex        =   12
            ToolTipText     =   "Random"
            Top             =   105
            Width           =   105
            _ExtentX        =   185
            _ExtentY        =   185
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Eq_SliderCtrl slider 
            Height          =   105
            Index           =   3
            Left            =   1740
            TabIndex        =   38
            ToolTipText     =   "Seek Bar"
            Top             =   390
            Width           =   1275
            _ExtentX        =   185
            _ExtentY        =   2249
            Value           =   60
            Position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl slider 
            Height          =   105
            Index           =   4
            Left            =   1740
            TabIndex        =   39
            ToolTipText     =   "Volume Bar"
            Top             =   120
            Width           =   975
            _ExtentX        =   185
            _ExtentY        =   1720
            Max             =   255
            Value           =   155
            Position        =   1
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   150
            Index           =   9
            Left            =   5520
            TabIndex        =   40
            ToolTipText     =   "Playlist"
            Top             =   330
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   150
            Index           =   10
            Left            =   6570
            TabIndex        =   41
            ToolTipText     =   "record"
            Top             =   330
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   150
            Index           =   11
            Left            =   6150
            TabIndex        =   42
            ToolTipText     =   "Crossfade"
            Top             =   330
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            Selected        =   -1  'True
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   150
            Index           =   12
            Left            =   5730
            TabIndex        =   43
            ToolTipText     =   "Repeat"
            Top             =   330
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   150
            Index           =   13
            Left            =   5940
            TabIndex        =   44
            ToolTipText     =   "Shuffle"
            Top             =   330
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
         Begin MMPlayerXProject.Button ButtonMini 
            Height          =   150
            Index           =   14
            Left            =   6360
            TabIndex        =   52
            ToolTipText     =   "Open"
            Top             =   330
            Width           =   150
            _ExtentX        =   265
            _ExtentY        =   265
            ButtonColor     =   65535
            MaskColor       =   16711935
            MousePointer    =   99
            Style           =   1
            UseMaskColor    =   -1  'True
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'CREDIT:  Raul Martinez Hernandez who started everything and gave me the way to move on
'Note: do not expect proper documentation, everything is in progress
'===========================================================================
'   Project : Neo Mp3 Player
'   Version : 2.0
'   Author  : Mahesh Kurmi
'   Email   : Mahesh_Kurmi2003@yahoo.com
'
'   '   You do NOT have rights to redistribute this code, in whole or in part
'   without my permission.  You also may not recompile the code and release
'   it as another program without my permission.  If you would like to modify
'   this code and distribute it in either as source code or as a compiled
'   program please contact me at [] before doing so.
'   I would appreciate being notified of any modifications even if you do not
'   intend to redistribute it...
'
'   Components:
'     - maheshmp3.dll  version 3.73 (in app path)
'     - Microsoft Common Dialog Control 6.0 (now not needed)
'     - Microsoft Windows Common Control 5.0
'     - Microsoft Windows Common Control 6.0'
Private currentEq As Integer
Private floatmode As Integer

Private blnLoaded As Boolean

'////// for spectrum displaY
Private band(50) As Double
Private ARRAYup(50) As Boolean
'//////////////
Private Type POINT
    X As Double
    Y As Double
End Type
Public activeSlider As Integer

'Private Const POINTS        As Long = 10

'=============================================================================
' WARNING: THIS PROGRAM USE SUBCLASSING, SO...
'          DO NOT PRESS THE STOP BUTTON IN VISUAL BASIC IDE!!!!
'=============================================================================

Dim ttDemo As New Tooltip
Public sSysTrayText As String
Dim PlayerIntro As Boolean
Dim TiempoIntro As Integer

Private randomCount As Integer
Dim PlayerMute As Boolean
'------------------------------------------------
Dim Slider_Selected As Integer    'formousewheel
'------------------------------------------------

'-----------------------------------------------
Public PlayerIsPlaying As String    '// WHERHER Player IS PLAYING OR NOT
Public VolumeNActuaL As Long    '// PRESENT VOLUME
'------------------------------------------------

Dim attach1 As Boolean    '//CHECK FOR ATTACHED PLAYLIST
Dim attach2 As Boolean    '//CHECK FOR ATTACHED visualisation

'// Variables FOR MOVING FORM
Dim bolDragMini As Boolean    '//DRAGGING IN MINI MODE
Dim StartDragX As Single
Dim StartDragY As Single
Dim rWorkArea As RECT
Dim mAttachedToRight As Boolean
Dim mAttachedToLeft As Boolean
Dim mAttachedToTop As Boolean
Dim mAttachedToBottom As Boolean
Dim mSnapDistance As Long
Dim bolTimeAct As Boolean

'// Spectrum
Dim arryPeaks(50) As Single
Dim arryWaitPeak(50) As String

'// Crossfade funcion VARIABLES
Dim lVol As Long
Dim lChannelOut As Long
Dim lChannelIn As Long
Dim iPresetIndex As Integer

Dim EqTempValues(10) As Double
Private cExToolTip As ExToolTip    '//-- ExTooltip Class
Private MMove As Boolean       '//-- Flag Used in ColorPicker for MouseMove Capture
Public cImage As c32bppDIB
Private tPic As IPictureDisp

'---------------------------------------------------------------------------------------
' Procedure : DrawEQ
' Author    : Mahesh
' UpDate    : 3/3/2011
' Purpose   : FUNCTION FOR DRAWING WAVY EQ CURVE
'---------------------------------------------------------------------------------------
'
Public Sub DrawEQ()
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim drwh As Single    'TO STORE PICTURES' WIDTH
    Dim drww As Single    'TO STORE PICTURES' WIDTH
    Dim drws As Long
    Dim p(10) As POINT

    Dim F As Long
    F = 5    'for smoothness of curve
    With PicEqAmp    'PICTURE WHERE EQ CURVE IS TO BE DRAWN
        PicEqAmp.AutoRedraw = True
        PicEqAmp.cls
        PicEqAmp.ScaleMode = 3
        drwh = .ScaleHeight - 1
        drww = .ScaleWidth
        PicEqAmp.BackColor = ColorBackGraph    'BACKGROUND COLOR SPECIFIED IN SKIN INI FILE
    End With

    drws = drww / 9    'TOTAL EQUALISER BANDS ARE 10 SO DIVIDE PICTURE WIDTH INTO 10 EQUAL PARTYS PARTS STARTING FROM 0

    For i = 0 To 9
        p(i).Y = (CDbl(drwh / 2)) - CDbl((frmMain.Eq_SliderCtrl(i).Value) * (drwh)) / 20    'TAKE Y CO-ORDINATE FROM EQ VALUE
    Next

    PicEqAmp.ForeColor = &H42585E
    PicEqAmp.Line (0, 0)-(0, drwh)    'DRAW LEFT VERTICAL LINE FOR 3D EFFECT
    PicEqAmp.Line (drww - 1, 0)-(drww - 1, drwh)    'DRAW RIGHT VERTICAL LINE FOR 3D EFFECT

    For i = 0 To 9
        PicEqAmp.Line (drws * i, (drwh / 2) - 3)-(drws * i, (drwh / 2) + 3)
        PicEqAmp.Line (drws * i + drws / 2, (drwh / 2) - 2)-(drws * i + drws / 2, (drwh / 2) + 2)
    Next

    PicEqAmp.Line (0, drwh / 2)-(drww, drwh / 2)    'DRAW MIDDLE HORIZONTAL LINE
    PicEqAmp.ForeColor = &H839EA7    'ANY RANDOM COLOR
    i = 1
    X = 0
    'Y = drwh / 2
    Y = p(0).Y    'CosineInterpolate(p(1).Y, p(2).Y, (X / F) / drws) 'INTERPOLATION FOR WAVY CURVE FROM LENEAR CURVE
    PicEqAmp.PSet (0, Y)

    Dim Xprev As Double
    Dim Yprev As Double
    Xprev = 0
    Yprev = Y
    For i = 0 To 9
        For X = 0 To drws * F    'DRAW CURVE IN STEPS
            Y = CosineInterpolate(p(i).Y, p(i + 1).Y, X / (F * drws))    ' - 0.5
            PicEqAmp.ForeColor = GetGradColor(drwh, Y, COLOR_END, COLOR_MIDDLE, COLOR_START)    'GET COLOR GRADIENT
            PicEqAmp.Line -(drws * i + (X / F), Y)    'DRAW CURVE IN STEPS
            Xprev = X + drws * i
            Yprev = Y
        Next
    Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : DrawAmplitudes
' Author    : Mahesh Kurmi
' Date      : 3/3/2011
' Purpose   : Draws wave oscillation(spectrum analyser)
'---------------------------------------------------------------------------------------
'
Public Sub DrawAmplitudes( _
       Data() As Integer, _
       picVis As PictureBox _
       )

    Dim dX As Long, dY As Long
    Dim X As Long, k As Long
    Dim dy2 As Long
    Dim dc0 As Long
    Dim lngPoints As Long
    Dim lngMaxAmpl As Long, lngAmpl As Long
    Dim dblAmpl As Double

    dX = picVis.ScaleWidth
    dY = picVis.ScaleHeight
    dy2 = dY \ 2
    dc0 = picVis.hDC

    picVis.ForeColor = tSpectrum.lBackColorScope

    For X = 2 To UBound(Data)
        lngAmpl = Abs(CLng(Data(X)))
        If Abs(lngAmpl) > lngMaxAmpl Then
            lngMaxAmpl = Abs(lngAmpl)
        End If
    Next
    ' lngMaxAmpl = 32767

    ' points per pixel
    lngPoints = (UBound(Data) / (1.3 * picVis.ScaleWidth)) / 2

    For X = 2 To picVis.ScaleWidth - 1
        ' average of some points
        dblAmpl = 0
        For k = k To k + lngPoints - 1
            dblAmpl = dblAmpl + Data(k)
        Next

        ' normalize points
        dblAmpl = (dblAmpl / lngPoints) / lngMaxAmpl
        If dblAmpl > 1 Then
            dblAmpl = 1
        ElseIf dblAmpl < -1 Then
            dblAmpl = -1
        End If

        ' draw a line to the new point
        LineTo dc0, X, -dblAmpl * (dy2) + dy2
    Next

    ' return to the middle
    LineTo dc0, X + 0, dy2
    LineTo dc0, X + 1, dy2
End Sub

'///////FUNCTION TO CONVERT LENEAR FUNCTION (Y) INTO COSINE FUNCTION (Y)
'---------------------------------------------------------------------------------------
' Procedure : CosineInterpolate
' Author    : Mahesh Kurmi
' Date      : 3/3/2011
' Purpose   : this function is used to draw sin^2(x) function between (x1,y1) and (x2,y2) and it gives value of f(x)
'             for given x in (x1,x2)
' HOW:        If assumes a function sin^2(x'*pi/2) for x' in (0,1) where x'=(x-x1)/(x2-x1) and this function joins
'             x1,y1) and (x2,y2) using smooth differentiable curve  sin^2(x).
' VARIBLES:
'             Y1=starting point  y coordinate or f(x1)
'             Y2=  Final y coordiante or f(x2)
'             mu=fractional value of x  or x'=(x-x1)/(x2-x1)
' RETURN:     This function gives the value of height of f(x) at any fractional x---------------------------------------------------------------------------------------
'
Private Function CosineInterpolate( _
        ByVal Y1 As Double, _
        ByVal Y2 As Double, _
        ByVal mu As Double _
        ) As Double

    Dim mu2 As Double
    mu2 = (1 - Cos(mu * Pi)) / 2    'same as sin^2(x)
    CosineInterpolate = (Y1 * (1 - mu2) + Y2 * mu2)
End Function

'---------------------------------------------------------------------------------------
' Procedure : Show_Message
' Author    : Mahesh Kurmi
' Date      : 3/3/2011
' Purpose   : FUNCTION TO DISPLAY MESSAGE IN SCROLLING TEXT FOR SMALL TIME INTERVAL
'---------------------------------------------------------------------------------------
'
Sub Show_Message(Message As String)
    If bMiniMask = True Then
        ScrollText(5).CaptionText = Message    'SHOW IN MINIMODE SCROLL DISPLAY
    Else
        ScrollText(1).CaptionText = Message
    End If
    Timer_Wait.Enabled = True    'TIMER TO RETAIN MESSAGE FOR 1 SECOND ...TIMER EVENT CHANGES THE DISPLAY ACFCORDINGLY

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Play_Crossfade
' Author    : Mahesh Kurmi
' Date      : 3/3/2011
' Purpose   : FOR CROSSFADING IN TRACKS
' HOW       : Decreases volume of prev channel and increases volume of new channel
' Returns   : -
' Parameters: -
'---------------------------------------------------------------------------------------
'
Sub Play_Crossfade()
    On Error Resume Next

    If (lCurrentChannel = 0) Then
        lCurrentChannel = 1: lChannelIn = 1: lChannelOut = 0    'IF NO STREAM IS OPENED OPEN ANOTHER STREAM IS SAVE ITS HANDLE TO DEFAULE ONE
    Else
        lCurrentChannel = 0: lChannelIn = 0: lChannelOut = 1
    End If

    Stream_Open sFileMainPlaying, FSOUND_NORMAL, lCurrentChannel, True, VolumeNActuaL
    Stream_SetBalance lCurrentChannel, slider(5).Value
    If PlayerMute = True Then Stream_SetMute lCurrentChannel, True

    lVol = VolumeNActuaL
    peCrossFadeType = CrossfadeNormal
    Timer_Crossfade.Interval = iCrossfadeTrack    'NO OF SECONDS FOR RETAINING CROSSFADE
    Timer_Crossfade.Enabled = True

End Sub

'/////FUNCTION FOR PLAY CLICK
'---------------------------------------------------------------------------------------
' Procedure : Play
' Author    : Mahesh Kurmi
' Date      : 3/3/2011
' Purpose   : Check which action is to be performed when play vutton is clicked
' HOW       :
' Returns   :
' Parameters:
'---------------------------------------------------------------------------------------
'
Sub Play()
    Timer_Player.Enabled = True
    If PlayerIntro = True Then Timer_Intro.Enabled = True: TiempoIntro = 0
    If PlayerIsPlaying = "pause" Then Pause_Play: Exit Sub
    If PlayerState = "true" Or PlayerState = "pause" Then
        frmTags.Stop_Player
    End If
    Start_Play
End Sub

'/////FUNCTION TO PLAY CURRENT FILE
'---------------------------------------------------------------------------------------
' Procedure : Start_Play
' Author    : Mahesh Kurmi
' Date      : 3/3/2011
' Purpose   : Starts play of track selected
'---------------------------------------------------------------------------------------
Sub Start_Play()
'On Error Resume Next
    If FileExists(sFileMainPlaying) = False Then Stop_Player: Stop_Draw_Spectrum: Next_Track: Exit Sub
    Load_File_Tags
    If iCrossfadeTrack <> 0 And bCrossFadeEnabled = True Then
        Play_Crossfade
    Else
        Stream_Stop lCurrentChannel
        Stream_Open sFileMainPlaying, FSOUND_NORMAL Or FSOUND_MPEGACCURATE, lCurrentChannel, True, VolumeNActuaL    'START NEW STREAM
        Stream_SetBalance lCurrentChannel, slider(5).Value
        'Debug.Print lCurrentChannel
    End If

    Timer_Player.Enabled = True
    PlayerIsPlaying = "true"
    Image_State_Rep    'DISPLAY SELECTED BUTTONS

    If bMiniMask = True Then    'SET MAX VALUE OF POSITION SLIDER
        slider(3).Max = CInt(Stream_GetDuration(lCurrentChannel))
    Else
        slider(0).Max = CInt(Stream_GetDuration(lCurrentChannel))
    End If
    Call frmLibrary.UpdatePlaycount(sFileMainPlaying, True)
    Call frmPLST.clist.ChangeEXTracklength(CurrentTrack_Index, Convert_Time_to_string(CInt(Stream_GetDuration(lCurrentChannel))))
    Call frmPLST.EnsureTrackVisible(CurrentTrack_Index)

    Exit Sub
error:
    Stop_Player
End Sub

'/////FUNCTION TO STOP PLAYER
Sub Stop_Player()
    On Error Resume Next
    PlayerIsPlaying = "false"
    If iCrossfadeStop = 0 Then Timer_Player.Enabled = False  'STOP RECORDER IF ON
    Image_State_Rep
    If PlayerIntro = True Then Timer_Intro.Enabled = False

    If iCrossfadeStop <> 0 And bCrossFadeEnabled = True Then
        'fade out
        lVol = VolumeNActuaL
        peCrossFadeType = FadeIn
        Timer_Crossfade.Interval = iCrossfadeStop
        Timer_Crossfade.Enabled = True
    Else
        Stream_Stop lCurrentChannel  '

        Stop_Draw_Spectrum
    End If

    If bMiniMask = True Then
        ScrollText(4).CaptionText = "00:00"
        slider(3).Value = 0
    Else
        ScrollText(0).CaptionText = "00:00"
        slider(0).Value = 0
    End If

End Sub

'//////////FUNCTION FOR PAUSE CLICK
Sub Pause_Play()
    Dim CurState As Long

    If PlayerIsPlaying = "false" Then Exit Sub

    CurState = Stream_GetState(lCurrentChannel)
    If CurState = 2 Then    'PAUSE

        If PlayerIntro = True Then Timer_Intro.Enabled = False
        PlayerIsPlaying = "pause"
        Image_State_Rep
        If iCrossfadeStop <> 0 Then
            '- Fade in -------------------------------------------------------
            lVol = VolumeNActuaL
            peCrossFadeType = FadeIn
            Timer_Crossfade.Interval = iCrossfadeStop
            Timer_Crossfade.Enabled = True
            '-----------------------------------------------------------------
        Else
            Stream_Pause lCurrentChannel
        End If

    Else
        PlayerIsPlaying = "true"
        Stream_Pause lCurrentChannel
        If iCrossfadeStop <> 0 And bCrossFadeEnabled = False Then
            Stream_SetVolume lCurrentChannel, 0
            '- Fade Out -------------------------------------------------------
            lVol = 0
            peCrossFadeType = FadeOut
            Timer_Crossfade.Interval = iCrossfadeStop
            Timer_Crossfade.Enabled = True
            '-----------------------------------------------------------------
        End If
        If PlayerIntro = True Then Timer_Intro.Enabled = True
    End If

    Image_State_Rep    'DISPLAY SELECTED BUTTONS

End Sub

'FUNCTION TO PLAY NEXT TRACK IN PLAYLIST
Sub Next_Track()
    If Button(8).Selected Then    'RANDOM IS ON
        CurrentTrack_Index = Random_Order_track(randomCount)
        If PlayerLoop = True And randomCount = frmPLST.clist.ItemCount - 1 Then randomCount = 0
        randomCount = randomCount + 1    'GO TO NEXT RANDOMISED TRACK FROM RANDOM ARRAY
        sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
        frmMain.PlayerIsPlaying = "true"
        frmPLST.ReinitializeList
        frmMain.Play
        Exit Sub
    End If

    If CurrentTrack_Index >= -1 And frmPLST.clist.ItemCount > 0 Then
        If CurrentTrack_Index = frmPLST.clist.ItemCount - 1 Then    ' FI TRACK IS LAST ONE MOVE TO FIRST IF LOOPING IS TRUE (REPEAT IS ON)
            If PlayerLoop = True Then
                CurrentTrack_Index = -1
            Else
                Exit Sub
            End If
        End If
        'check if currenttrack is not the last track
        If CurrentTrack_Index < (frmPLST.clist.ItemCount - 1) Then CurrentTrack_Index = CurrentTrack_Index + 1    'GOTO NEXT TRACK NOW LIST IS BEING REPEATED
        sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)    'GET CURRENT TRACKPATH FROM PLAYLIST'S LIST OBJECT
        frmMain.PlayerIsPlaying = "true"
        frmPLST.ReinitializeList    'REDRAW PLAYLIST
        DoEvents
        frmMain.Play
    End If

End Sub

'////FUNCTION TO PLAY PREVIOSLY PLAYED TRACK IN PLAYLIST

Sub Previous_Track()

    If Button(8).Selected Then    'RANDOM IS ON
'Static randomCount As Integer
        CurrentTrack_Index = Random_Order_track(randomCount)
        randomCount = randomCount - 1    'GO TO NEXT RANDOMISED TRACK FROM RANDOM ARRAY
        If randomCount < 0 Then randomCount = 0
        sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
        frmMain.PlayerIsPlaying = "true"
        frmPLST.ReinitializeList
        frmMain.Play
        Exit Sub
    End If

    If CurrentTrack_Index < frmPLST.clist.ItemCount And CurrentTrack_Index > 0 Then
        CurrentTrack_Index = CurrentTrack_Index - 1
        sFileMainPlaying = frmPLST.clist.exItem(CurrentTrack_Index)
        frmMain.PlayerIsPlaying = "true"
        frmPLST.ReinitializeList
        DoEvents
        frmMain.Play
    End If

End Sub

'////FORWARD BY 5 SECONDS
Sub Five_Seg_Backward()
    On Error Resume Next
    If PlayerIsPlaying = "false" Then Exit Sub

    Dim CurPos As Long
    'If ListRep.ListCount = 0 Or PlayerIsPlaying <> "true" Then Exit Sub
    If iPosScrollChange = 0 Then iPosScrollChange = 5
    CurPos = Stream_GetPosition(lCurrentChannel)
    CurPos = CurPos - iPosScrollChange
    If CurPos < 0 Then CurPos = 0
    Stream_SetPosition lCurrentChannel, CurPos
    Show_Message Convert_Time_to_string(CurPos) + "/" + Convert_Time_to_string(CurPos) + "(" + CStr(CInt((CurPos / slider(1).Max) * 100)) + "%)"

End Sub

'///INTRO SONG OF 10 SECONDS
Sub Intro()

    If PlayerIntro = False Then
        'ACTIVATE INTRO
        Button(5).Selected = True
        PlayerIntro = True
        TiempoIntro = 0
        Timer_Intro.Enabled = True
        frmPopUp.mnuIntro.Checked = True
        Show_Message "Intro ON"
    Else
        'DEACTIVATE INTRO
        Button(5).Selected = False
        PlayerIntro = False
        Timer_Intro.Enabled = False
        frmPopUp.mnuIntro.Checked = False
        Show_Message "Intro OFF"
    End If

End Sub

'///////MUTE THE STREAM
Sub Player_Mute()
    On Error Resume Next
    If PlayerMute = False Then
        '--activATE silencE--------------------
        Button(6).Selected = True
        PlayerMute = True
        Stream_SetMute lCurrentChannel, True
        frmPopUp.mnuSilencio.Checked = True
        Show_Message "Mute ON"
    Else
        'DeactivATE SILENCE-----------------------------
        Button(6).Selected = False
        PlayerMute = False
        Stream_SetMute lCurrentChannel, False
        frmPopUp.mnuSilencio.Checked = False
        Show_Message "Mute OFF"
    End If
End Sub

'//////LOOP CLICK
Sub Player_Repeat(bLoop As Boolean)
    If bLoop = True Then
        '---ActivATE loop -----------------------------
        Button(7).Selected = True
        ButtonMini(12).Selected = True
        frmPopUp.mnuRepeatTrack.Checked = True
        Show_Message "Repeat ON"
        PicRepeat.PaintPicture PicConfigInfo, 0, 0, PicRepeat.Width, PicRepeat.Height, PicShuffle.Width, PicRepeat.Height, PicRepeat.Width, PicRepeat.Height
        PicRepeat.Picture = PicRepeat.Image

    Else
        '--- DescativATE loop ---------------------------
        Button(7).Selected = False
        ButtonMini(12).Selected = False
        frmPopUp.mnuRepeatTrack.Checked = False
        Show_Message "Repeat OFF"
        PicRepeat.PaintPicture PicConfigInfo, 0, 0, PicRepeat.Width, PicRepeat.Height, PicShuffle.Width, 0, PicRepeat.Width, PicRepeat.Height
        PicRepeat.Picture = PicRepeat.Image
    End If
End Sub

'//////LOAD FILE TAGS FOR ARTIST DESPLAY IN SCROLL TEXT
Sub Load_File_Tags()
    On Error Resume Next
    Dim sFullpath As String, sFilename As String, sFileEx As String
    Dim aFile() As String

    ' load tags
    tCurrentID3.Read_MPEGInfo = True
    tCurrentID3.Read_File_Tags sFileMainPlaying

    ScrollText(2).CaptionText = tCurrentID3.MPEG_Bit_Rate
    ScrollText(3).CaptionText = CStr(Left(tCurrentID3.MPEG_Frequency, 2))
    sFullpath = sFileMainPlaying
    aFile = Split(sFullpath, "\", , vbTextCompare)
    sFilename = aFile(UBound(aFile))
    sFileEx = Right(sFullpath, 3)

    sSysTrayText = tCurrentID3.Title & " - " & tCurrentID3.Artist & " - NoMP3player"
    If OpcionesMusic.TaskBar = True Then frmPopUp.Caption = sSysTrayText
    'If OpcionesMusic.SysTray = True Then CambiarIcono Text1.hWnd, Me.Icon.handle, sSysTrayText



    '*****************************************************************************
    '//////UPDATE SCROLL TEXT
    sTextScroll = ""
    '// Song Name
    If Trim(tCurrentID3.Title) = "" Then tCurrentID3.Title = GetFileTitle(sFileMainPlaying)
    sTextScroll = Replace(sFormatScroll, "%S", tCurrentID3.Title)
    '// Artist
    sTextScroll = Replace(sTextScroll, "%A", tCurrentID3.Artist)
    '// Album
    sTextScroll = Replace(sTextScroll, "%B", tCurrentID3.Album)
    '// Year
    sTextScroll = Replace(sTextScroll, "%Y", tCurrentID3.Year)
    '// Genre
    sTextScroll = Replace(sTextScroll, "%G", tCurrentID3.Genre)
    '// Time
    sTextScroll = Replace(sTextScroll, "%T", tCurrentID3.MPEG_DurationTime)
    '// File Name
    sTextScroll = Replace(sTextScroll, "%N", sFilename)
    '// Time
    sTextScroll = Replace(sTextScroll, "%P", sFullpath)
    '// File extencion
    sTextScroll = Replace(sTextScroll, "%F", sFileEx)
    If sTextScroll = sFormatScroll Then sTextScroll = tCurrentID3.Title

    If bMiniMask = True Then
        ScrollText(5).CaptionText = sTextScroll
        ScrollText(5).ToolTipText = sTextScroll
    Else
        ScrollText(1).CaptionText = sTextScroll
        ScrollText(1).ToolTipText = sTextScroll
    End If

    If boolVisShow Then FrmVisualisation.ScrollText1.CaptionText = sTextScroll
    '*****************************************************************************




    '*****************************************************************************
    '//////UPDATE PLAYLIST TRACK ALSO
    sTextScroll = ""
    '// Song Name
    sTextScroll = Replace(sFormatPlayList, "%S", tCurrentID3.Title)
    '// Artist
    sTextScroll = Replace(sTextScroll, "%A", tCurrentID3.Artist)
    '// Album
    sTextScroll = Replace(sTextScroll, "%B", tCurrentID3.Album)
    '// Year
    sTextScroll = Replace(sTextScroll, "%Y", tCurrentID3.Year)
    '// Genre
    sTextScroll = Replace(sTextScroll, "%G", tCurrentID3.Genre)
    '// Time
    sTextScroll = Replace(sTextScroll, "%T", tCurrentID3.MPEG_DurationTime)
    '// File Name
    sTextScroll = Replace(sTextScroll, "%N", sFilename)
    '// Time
    sTextScroll = Replace(sTextScroll, "%P", sFullpath)
    '// File extencion
    sTextScroll = Replace(sTextScroll, "%F", sFileEx)

    If sTextScroll = sFormatPlayList Then sTextScroll = tCurrentID3.Title

    Call frmPLST.clist.ChangeItem(CurrentTrack_Index, UpperCase_Firstletter(sTextScroll))
    '*****************************************************************************



    '// UPDATE LYRICS IF ANY
    '*****************************************************************************
    If tCurrentID3.Has_Lyrics3_Tag = True And Trim(tCurrentID3.Lyrics) <> "" Then
        'Show_Lyrics Trim(tCurrentID3.Lyrics)
    Else
        'LyricsRef.Clear
    End If
    '*****************************************************************************



    '// UPDATE TOOLTIP of Song INFO
    '*****************************************************************************
    Dim k As Long
    cImage.DestroyDIB
    If ReadAlbumArt(sFileMainPlaying, 1, tPic, k) Then
        Call cImage.LoadPicture_StdPicture(tPic)
        'Load Cimage to Picturebox
        cImage.ScaleImage 120, 132, 120, 132, scaleDownAsNeeded
    End If
    '// prepare tooltiptext
    Show_ToolTipText
    '*****************************************************************************


End Sub

'//////// TOOLTIP ON MOUSE MOVE FOR SONG INFORMATION DISPLAY
Sub Show_ToolTipText()
' On Error GoTo hell
    Exit Sub
    ttDemo.ForeColor = NormalText_Forecolor
    ttDemo.BackColor = frmPLST.picList.BackColor    'rgb(0,0,0) then it paints white =>>don't know why???
    ttDemo.TipText = "Track  : " & Format$(str(CurrentTrack_Index + 1), "00") + "/" + Format$(str(frmPLST.clist.ItemCount), "00") & vbCrLf & _
                     "Titl    e: " & tCurrentID3.Title & vbCrLf & _
                     "Length: " & tCurrentID3.MPEG_DurationTime & vbCrLf & _
                     "Artist  : " & tCurrentID3.Artist & vbCrLf & _
                     "Album : " & tCurrentID3.Album & vbCrLf & _
                     "Year   : " & tCurrentID3.Year & vbCrLf & _
                     "Genre : " & tCurrentID3.Genre & vbCrLf & _
                     "Path   : " & sFileMainPlaying
    If bMiniMask = True Then
        'usercontrol isn't showing tooltip: i don't know why? hence i used tooltip of picturebox of user control
        ttDemo.ParentControl = ScrollText(5).picHandle    'picMiniMode ' ScrollText(5)
    Else
        ttDemo.ParentControl = ScrollText(1).picHandle    'ListRep ' picNormalMode
    End If
    ttDemo.Title = "MaheshMP3 Player"
    ttDemo.Icon = TTIconInfo
    ttDemo.Style = TTBalloon
    ttDemo.Create
hell:
End Sub

'//////// DISPLAY BSELECTED BUTTONS
Sub Image_State_Rep()
    On Error Resume Next
    If bMiniMask = False Then    '// NORMAL MODE
        Button(1).Selected = False
        Button(2).Selected = False
        Button(3).Selected = False
        Select Case PlayerIsPlaying
        Case "true"  'PLAY
            Button(1).Selected = True
            PicInfo.PaintPicture PicConfigInfo, 0, 0, PicInfo.Width, PicInfo.Height, 0, 2 * PicShuffle.Height, PicInfo.Width, PicInfo.Height
            PicInfo.Picture = PicInfo.Image
        Case "false"    'STOP
            Button(3).Selected = True
            PicInfo.PaintPicture PicConfigInfo, 0, 0, PicInfo.Width, PicInfo.Height, PicInfo.Width, 2 * PicShuffle.Height, PicInfo.Width, PicInfo.Height
            PicInfo.Picture = PicInfo.Image
        Case "pause"    'PausE
            Button(2).Selected = True
            PicInfo.PaintPicture PicConfigInfo, 0, 0, PicInfo.Width, PicInfo.Height, 2 * PicInfo.Width, 2 * PicShuffle.Height, PicInfo.Width, PicInfo.Height
            PicInfo.Picture = PicInfo.Image
        End Select
    Else    '// MINI MODE
        ButtonMini(1).Selected = False
        ButtonMini(2).Selected = False
        ButtonMini(3).Selected = False
        Select Case PlayerIsPlaying
        Case "true"
            ButtonMini(1).Selected = True
        Case "false"
            ButtonMini(3).Selected = True
        Case "pause"
            ButtonMini(2).Selected = True
        End Select
    End If

End Sub

'//////DRAW SPECTRUM DISPLAY
Public Sub Draw_Spectrum(PicSpectrum As PictureBox)
    Dim X1 As Single, Y1 As Single
    Dim X2 As Single, Y2 As Single
    Dim iPeak As Single
    Dim iSleep As Integer
    Dim i As Long
    Dim j As Long
    'Dim sngRealOut(FFT_SAMPLES - 1) As Single
    On Error Resume Next
    Dim k, lngcolor
    Dim sngSpectrumData() As Single

    If floatmode = 4 Then Stop_Draw_Spectrum
    Dim sngBand As Single
    Dim rcBand As RECT
    Static m As Boolean
    If m = False Then

        For i = 0 To 36    'STORE PEAK VALUES IN AN ARRAY
            arryPeaks(i) = PicSpectrum.Height
        Next
        m = True
    End If

    With PicSpectrum
        .cls

        Dim ints(1023) As Single
        Dim iint(1023) As Integer
        Dim lpspectrum As Long

        If bLoadingSkin Or bLoading Then Exit Sub
        '////// visualisation display
        Static aaa(1023) As Integer
        If boolVisShow And boolVisLoaded Then
            lpspectrum = FSOUND_DSP_GetSpectrum    'get reference of pointer
            CopyMemory ints(0), ByVal lpspectrum, 1024 * 4    'convert long into integer  using pointers
            For i = 0 To 1023 Step 1
                iint(i) = ints(i) * 1000000
                'ints(i + 1) = sngSpectrumData(i) * 1000000
            Next
            FrmVisualisation.ScopeBuff.cls
            FrmVisualisation.PicSpectrum.cls
            If boolOpluginLoaded Then
                oPlugIn.drawVis FrmVisualisation.PicSpectrum.hDC, iint(), FrmVisualisation.ScopeBuff.Height, FrmVisualisation.ScopeBuff.Width
            Else
                FrmVisualisation.drawVis FrmVisualisation.PicSpectrum.hDC, iint(), FrmVisualisation.ScopeBuff.Height, FrmVisualisation.ScopeBuff.Width
            End If
            'FrmVisualisation.PicSpectrum.Picture = FrmVisualisation.ScopeBuff.Image

        End If
        '///// end of vis display


        If frmPopUp.mnuSpecOsc.Checked = True Then    ' scope
            ReDim sngSpectrumData(50) As Single
            Spectrum_GetData lCurrentChannel, sngSpectrumData()

            With PicSpectrum
                .CurrentY = .ScaleHeight / 2
                .CurrentX = 0
                For i = 0 To tSpectrum.iLinesScope
                    X1 = i * (.ScaleWidth / tSpectrum.iLinesScope)
                    X2 = X1 + (.ScaleWidth / tSpectrum.iLinesScope)
                    Y1 = .ScaleHeight / 2
                    Y2 = (sngSpectrumData(i) * Y1)
                    PicSpectrum.DrawWidth = 1
                    PicSpectrum.Line Step(0, 0)-(X1 + ((X2 - X1) / 3), Y1 - Y2), tSpectrum.lBackColorScope
                    PicSpectrum.Line Step(0, 0)-(X1 + (((X2 - X1) / 3) * 2), Y1 + Y2), tSpectrum.lBackColorScope
                    PicSpectrum.Line Step(0, 0)-(X2, Y1), tSpectrum.lBackColorScope
                Next i
            End With
            Exit Sub

        ElseIf frmPopUp.mnuSpecNone.Checked = True Then
            lpspectrum = FSOUND_DSP_GetSpectrum    'get reference of pointer
            CopyMemory ints(0), ByVal lpspectrum, 512 * 4    'convert long into integer  using pointers
            For i = 0 To 511 Step 1
                iint(i) = ints(i) * 1000000
                'iint(823 - i) = iint(i)
                'ints(i + 1) = sngSpectrumData(i) * 1000000
            Next
            Call DrawAmplitudes(iint, PicSpectrum): Exit Sub
        End If



        ReDim sngSpectrumData(512) As Single
        Spectrum_GetData lCurrentChannel, sngSpectrumData()
        For i = 0 To FFT_BANDS
            band(i) = band(i) - FFT_BANDLOWER
            If band(i) < 0 Then band(i) = 0
        Next

        For i = 0 To 128
            X1 = i * (DRW_BARWIDTH + DRW_BARSPACE)  '(.ScaleWidth / tSpectrum.iBars)
            X2 = X1 + DRW_BARWIDTH  '(CInt(.ScaleWidth / tSpectrum.iBars) - 1)
            Y1 = .ScaleHeight
            If sngSpectrumData(i) > FFT_MAXAMPLITUDE Then
                sngSpectrumData(i) = FFT_MAXAMPLITUDE
            End If
            sngSpectrumData(i) = sngSpectrumData(i) / FFT_MAXAMPLITUDE
        Next

        j = FFT_STARTINDEX

        For i = 0 To FFT_BANDS
            Dim T As Integer
            If i >= FFT_BANDS * 3 / 5# Then
                T = 2
            Else
                T = 4
            End If
            ' average for the current band
            For j = j To j + T
                sngBand = sngBand + sngSpectrumData(j)
            Next
            ' boost frequencies in the middle with a hanning window,
            ' because they got less power then the low ones
            sngBand = (sngBand * (Hanning(i + 3, 25 + 3) + 1)) / T

            If band(i) < sngBand Then band(i) = sngBand
            If band(i) > 1 Then band(i) = 1
            ' skip some bands
            j = j + FFT_BANDSPACE
        Next


        For i = 0 To FFT_BANDS

            With rcBand
                .Right = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARWIDTH + DRW_BARXOFF
                .Left = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARXOFF
                .Top = maxX(DRW_BARYOFF, Min(PicSpectrum.ScaleHeight - 1, PicSpectrum.ScaleHeight - (PicSpectrum.ScaleHeight * band(i))) - DRW_BARYOFF)    ' - 1)
                .Bottom = PicSpectrum.ScaleHeight - DRW_BARYOFF
            End With
            If arryPeaks(i) >= rcBand.Top Then
                arryPeaks(i) = rcBand.Top
                ARRAYup(i) = True
                aaa(i) = 10
                'arryWaitPeak(i) = time
                'If floatMode <> 1 Then arryWaitPeak(i) = 0
            End If

            For k = 1 To PicSpectrum.ScaleHeight - rcBand.Top - 1 Step Bar_gap
                lngcolor = GetGradColor(1, k / PicSpectrum.ScaleHeight, COLOR_1, COLOR_2, COLOR_3)
                PicSpectrum.Line (rcBand.Left, PicSpectrum.ScaleHeight - k)-(rcBand.Right, PicSpectrum.ScaleHeight - k), lngcolor    'RGB(intRed, intGreen, intBlue) ', BF
            Next

            If floatmode = 3 Then GoTo continue_for:

            If arryPeaks(i) < 0 Then
                arryPeaks(i) = 0
                If floatmode = 2 Then arryPeaks(i) = rcBand.Top
                ARRAYup(i) = False
            End If
            iPeak = arryPeaks(i)

            If Peak_Color = 0 Then    'ini setting is not there
                lngcolor = GetGradColor(1, 1 - (iPeak) / PicSpectrum.ScaleHeight, COLOR_1, COLOR_2, COLOR_3)
            Else
                lngcolor = Peak_Color
            End If
            PicSpectrum.Line (rcBand.Left, iPeak - 1)-(rcBand.Right, iPeak - 1), lngcolor    ', BF


            If arryWaitPeak(i) <> "" Then
                iSleep = DateDiff("s", arryWaitPeak(i), Time)
            End If

            aaa(i) = aaa(i) - 1
            If aaa(i) <= 0 Or floatmode = 1 Then
                aaa(i) = 0
                Select Case floatmode
                Case 0:
                    arryPeaks(i) = arryPeaks(i) + DRW_PEAKFALL
                Case 1:
                    If ARRAYup(i) = True Then
                        arryPeaks(i) = arryPeaks(i) - DRW_PEAKFALL
                    Else
                        arryPeaks(i) = arryPeaks(i) + DRW_PEAKFALL
                    End If
                Case 2:
                    arryPeaks(i) = arryPeaks(i) - DRW_PEAKFALL
                End Select
            End If
continue_for:
        Next

    End With


End Sub

' DISPLAY STOPPED SPECTRUM STATE
Public Sub Stop_Draw_Spectrum()
    Dim iPeak As Single
    Dim i As Long
    Dim Left, Right, lngcolor
    PicSpectrum.cls
    PicSpectrumMini.cls
    For i = 0 To 36
        arryPeaks(i) = PicSpectrum.Height - 1
    Next

    For i = 0 To FFT_BANDS - 1
        iPeak = arryPeaks(i)
        Right = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARWIDTH + DRW_BARXOFF
        Left = i * (DRW_BARWIDTH + DRW_BARSPACE) + DRW_BARXOFF
        lngcolor = GetGradColor(1, 1 - (iPeak) / PicSpectrum.ScaleHeight, COLOR_1, COLOR_2, COLOR_3)
        PicSpectrum.Line (Left, iPeak - 1)-(Right, iPeak - 1), lngcolor    ', BF
        'PicSpectrumMini.Line (Left, iPeak - 1)-(Right, iPeak - 1), lngcolor ', BF
    Next

End Sub

Private Sub BALANCE_Change(Value As Long)
    Call Slider_Change(5, Value)
End Sub

Public Sub Button_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0: Previous_Track
    Case 1: Play
    Case 2: Pause_Play
    Case 3: Stop_Player
    Case 4: Next_Track
    Case 5: Intro
    Case 6: Player_Mute
    Case 7: frmPopUp.mnuRepeatTrack_Click
    Case 8: Button(8).Selected = Not Button(8).Selected
        ButtonMini(13).Selected = Button(8).Selected
        frmPopUp.mnuOrdenAleatorio.Checked = Button(8).Selected
        If Button(8).Selected Then
            Show_Message "SHUFFLE ON"
            PicShuffle.PaintPicture PicConfigInfo, 0, 0, PicShuffle.Width, PicShuffle.Height, 0, 0, PicShuffle.Width, PicShuffle.Height
            PicShuffle.Picture = PicShuffle.Image
        Else
            Show_Message "SHUFFLE OFF"
            PicShuffle.PaintPicture PicConfigInfo, 0, 0, PicShuffle.Width, PicShuffle.Height, 0, PicShuffle.Height, PicShuffle.Width, PicShuffle.Height
            PicShuffle.Picture = PicShuffle.Image
        End If
        RANDOM_track

    Case 10:    'frmPopUp.mnuDSP_Click
        Call FSOUND_SetSurround(lCurrentChannel, 90)
    Case 9: Setpreset (1)
    Case 11: Setpreset (-1)
    Case 12: Me.PopupMenu frmPopUp.mnuMenuPrincipal
    Case 13: Minimize_Me
    Case 14: Change_Mask True, True: frmMain.Image_State_Rep
    Case 15:    ' Call frmPopUp.vkSysTray1.RemoveFromTray(1) 'i will use it later
        Unload frmMain
    Case 16: Button(16).Selected = Not Button(16).Selected
        Dim i As Integer
        FX_Disable  'Disable all FX
        If (Button(16).Selected = True) Then
            FX_Enable 7
            For i = 0 To 9 Step 1
                FX_SetEQ CLng(i), CLng(Eq_SliderCtrl(i).Value)
            Next
        End If

    Case 17: Button(17).Selected = Not Button(17).Selected    'Eq FreeMovement On/OFF
        bEQFreeMove = Not bEQFreeMove
        'If bEQFreeMove Then
        'Call FSOUND_SetSurround(lCurrentChannel, CInt(bEQFreeMove) * 100)
        If bEQFreeMove Then
            FSOUND_SetSpeakerMode (FSOUND_SPEAKERMODE_SURROUND)
        Else
            FSOUND_SetSpeakerMode (FSOUND_SPEAKERMODE_MONO)
        End If

    Case 18:
        If boolDspShow = True Then    'Frmdsp already loaded
            frmDSP.Show
        Else
            frmPopUp.mnuDSP_Click    'load frmdsp
        End If
        frmDSP.vkBtn_Click (0)    'bring equalizer tab to the front

    Case 22: Button(22).Selected = Not Button(22).Selected    'unused button

    Case 23: frmPopUp.mnuPreferences_Click

    Case 24: frmMain.Button(24).Selected = Not frmMain.Button(24).Selected
        If frmMain.Visible = True Then frmMain.EqBackpic.Visible = frmMain.Button(24).Selected
        frmPopUp.mnuShowEqualizer.Checked = frmMain.Button(24).Selected
        Call CheckMenu(10, frmPopUp.mnuShowEqualizer.Checked)
        EQ_in_out    'Show/hide equaliser 15/3/2008

    Case 25: Call frmPLST.ADD_MULTIPLE_FILES(True)    'true is set if new list is to be constructed
    Case 26:    'show/hide playlist
        Button(Index).Selected = Not Button(Index).Selected
        frmPLST.Visible = Button(Index).Selected
        ButtonMini(9).Selected = Button(Index).Selected
        frmPopUp.mnuShowPlst.Checked = frmPLST.Visible
        Call CheckMenu(11, frmPLST.Visible)
    End Select
    Timer_Wait.Enabled = True    'for Scroll display

End Sub

'//EQUALISERS ARE TO BE RESET
Private Sub Eq_SliderCtrl_Change(Index As Integer, Value As Long)
    Dim lngX As Long
    'On Error GoTo error:
    Dim i As Integer
    lngX = CLng(Index)
    Dim eQval(10) As Long
    'DoEvents

    If currentEq <> Index And activeSlider = -1 Then Exit Sub

    'The change event  calls itself so all onditions ae imposed for out of stack error
    If bEQFreeMove Then
        Dim iChange As Double
        iChange = 0
        iChange = CDbl(Eq_SliderCtrl(Index).Value - EqTempValues(Index))
        For i = 0 To 9
            If iChange = 0 Or activeSlider <> -1 Then Exit For
            Dim j As Double
            j = i - Index
            j = 1# + (2 * (j / 2) ^ 4)
            j = iChange / j
            eQval(i) = EqTempValues(i) + j    'EqTempValues(Index) + j '
            If i <> Index Then
                If eQval(i) > 10 Then
                    Eq_SliderCtrl(i).Value = 10
                ElseIf eQval(i) < -10 Then
                    Eq_SliderCtrl(i).Value = -10
                Else
                    Eq_SliderCtrl(i).Value = eQval(i)
                End If
            End If
            If Button(16).Selected Then FX_SetEQ CLng(i), CLng(eQval(i))    'FX_SetEQ Index, -sldEQ(Value).Value
        Next
    End If
    'If eq is enabled then set DFX
    If Button(16).Selected Then FX_SetEQ CLng(Index), CLng(Value)    'FX_SetEQ Index, -sldEQ(Value).Value
    ScrollText(1).CaptionText = "EQ BAND" + "." & (Index + 1) & " " & str(Eq_SliderCtrl(Index).Value) & " dB"
    DrawEQ
error:
End Sub

Private Sub Eq_SliderCtrl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To 9
        EqTempValues(i) = Eq_SliderCtrl(i).Value
    Next
    activeSlider = -1
    currentEq = Index
End Sub

'dISPLAY EQ VALUES IN SCROLL TEXT
Private Sub Eq_SliderCtrl_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    currentEq = -1
    Timer_Wait.Enabled = True
End Sub

Private Sub EqBackpic_GotFocus()
    Me.SetFocus
End Sub

'MOVE FORM
Private Sub EqBackpic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    activeSlider = -1
    currentEq = -1
End Sub

Private Sub EqBackpic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then

        Dim i
        For i = 0 To 9
            If Eq_SliderCtrl(i).Left < X And Eq_SliderCtrl(i).Left + Eq_SliderCtrl(i).Width > X Then
                activeSlider = i
                Exit For
            End If
        Next
        If activeSlider = -1 Then Exit Sub
        If Y <= Eq_SliderCtrl(0).Top And Eq_SliderCtrl(activeSlider).Value <> 10 Then Eq_SliderCtrl(activeSlider).Value = 10: Exit Sub
        Eq_SliderCtrl(activeSlider).Value = -10 + Min(20, ((Eq_SliderCtrl(0).Height + Eq_SliderCtrl(0).Top - Y) * 21#) / Eq_SliderCtrl(0).Height)
    End If
End Sub

Private Sub EqBackpic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bolDragMini = False
    activeSlider = -1
End Sub

Public Sub ActivateMe()
'On Error Resume Next
    Dim i
    If bLoadingSkin Then
        picNormalMode.Refresh
        EqBackpic.Refresh
    End If
    PicActivate.Visible = False
    For i = 12 To 15
        Button(i).Selected = False
    Next
End Sub

Public Sub DeactivateMe()
    PicActivate.Visible = True
    Dim i
    For i = 12 To 15
        Button(i).Selected = True
    Next
End Sub

Private Sub Form_Activate()
    ActivateMe
End Sub

Private Sub Form_Initialize()
    InitCommonControls                   '//-- Registers and initializes the common control window classes.
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    'index = Slider_Selected
    If KeyCode = &H28 Then    ' - volumen keycode=40
        VolumeNActuaL = VolumeNActuaL - 4
        If VolumeNActuaL < 0 Then VolumeNActuaL = 0
        If VolumeNActuaL > 255 Then VolumeNActuaL = 255
        slider(1).Value = VolumeNActuaL
        slider(4).Value = VolumeNActuaL
        Slider_Change 1, VolumeNActuaL
        Slider_Change 4, VolumeNActuaL
    End If
    If KeyCode = &H26 Then    ' + volumen'keycode=38
        VolumeNActuaL = VolumeNActuaL + 4
        If VolumeNActuaL < 0 Then VolumeNActuaL = 0
        If VolumeNActuaL > 255 Then VolumeNActuaL = 255
        slider(1).Value = VolumeNActuaL
        slider(4).Value = VolumeNActuaL
        Slider_Change 1, VolumeNActuaL
        Slider_Change 4, VolumeNActuaL
    End If

    If KeyCode = &H25 Then Five_Seg_Backward  'A Atras 5 seg  keycode=37
    If KeyCode = &H27 Or KeyCode = 100 Then Five_Seg_Forward    'D Adelante 5 seg keycode=39

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Form_Unload(Cancel)

End Sub

Private Sub Form_Resize()
    If bLoading Or bLoadingSkin Then Exit Sub
    If Me.WindowState = vbMinimized Then
        If boolOptionsShow = True And boolOptionsLoaded Then frmOptions.Hide
        If boolTagsShow = True Then frmTags.Hide
        If boolDspShow = True And boolDspLoaded Then frmDSP.Hide
        frmPLST.Hide
        'frmPopUp.Show
        If boolVisShow = True Then FrmVisualisation.Hide
        If boolMediaLibraryShow = True Then frmLibrary.Hide
    Else
        If boolVisShow = True And boolVisLoaded Then FrmVisualisation.Visible = True
        frmPLST.Visible = frmMain.Button(26).Selected
        'frmPopUp.Hide
        If boolDspShow = True And boolDspLoaded Then frmDSP.Visible = True
        If boolOptionsShow = True And boolOptionsLoaded Then frmOptions.Visible = True
        'frmMain.Show 'to set focus on frmmain Show method sets focus on control to be shown
        If boolMediaLibraryShow = True Then frmLibrary.Show
    End If
End Sub

Private Sub Form_Terminate()
    Call Form_Unload(0)
End Sub

Private Sub PicActivate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.SetFocus
End Sub

Private Sub PicaTop_Click()
    frmPopUp.Atop_Click
End Sub

Private Sub PicCrossfade_Click()
    bCrossFadeEnabled = Not bCrossFadeEnabled
    EnableCrossfade (bCrossFadeEnabled)
End Sub

Private Sub picNormalMode_GotFocus()
    On Error Resume Next
    Me.SetFocus
End Sub

Private Sub picNormalMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If MMove Then MMove = False
    FormDrag_Move X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
End Sub

Private Sub picNormalMode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bolDragMini = False
End Sub

Private Sub PicRepeat_Click()
    frmPopUp.mnuRepeatTrack_Click
End Sub

Private Sub PicShuffle_Click()
    Call Button_Click(8)
End Sub

Private Sub picSpectrum_Click()
    If floatmode = 4 Then floatmode = -1
    floatmode = floatmode + 1
End Sub

Private Sub picSpectrum_DblClick()
    With frmPopUp
        If .mnuSpecNone.Checked = True Then
            .mnuSpecBars_Click
        ElseIf .mnuSpecBars.Checked = True Then
            .mnuSpecOsc_Click
        ElseIf .mnuSpecOsc.Checked = True Then
            .mnuSpecNone_Click
        End If

    End With
    floatmode = 0
End Sub

Private Sub PicSpectrum_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' If Button = vbRightButton Then PopupMenu Me.aaa

End Sub

Private Sub PicSpectrumMini_Click()
    Call picSpectrum_Click
End Sub

Private Sub PicSpectrumMini_DblClick()
    Call picSpectrum_DblClick
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub ScrollText_DblClick(Index As Integer)
'// show diferent curent time
    On Error Resume Next
    If sFileMainPlaying = "" Then Exit Sub
    If Index = 0 Or Index = 4 Then
        bolTimeAct = Not bolTimeAct
        'ScrollText(Index).Left = ScrollText(Index).Left - ScrollText(Index).
    End If
    '// stop scroll
    If (Index = 1 Or Index = 5) And FileExists(sFileMainPlaying) Then
        'frmTags.c
        'frmTags.fileTags.Clear
        frmTags.Show
        If cGetInputState() <> 0 Then DoEvents
        frmTags.vkFiletags.Clear
        frmTags.listRef.ListItems.Clear
        frmTags.Load_Tags sFileMainPlaying
        frmTags.vkFiletags.Selected(1) = True    'index starts from 1 in vkListbox instead of zero
        frmTags.Show_tags (1)
        ' frmTags.fileTags.Selected(0) = True
    End If


End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Public Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Debug.Print Shift
    If Shift = 0 Then
        If KeyCode = 66 Then Previous_Track    'Z
        If KeyCode = 80 Then Play    'X
        If KeyCode = 32 Then Pause_Play    'P
        If KeyCode = 86 Then Stop_Player    'V
        If KeyCode = 78 Then Next_Track    'B
        If KeyCode = 73 Then Intro    'I Intro 10 seg
        If KeyCode = 82 Then frmPopUp.mnuRepeatTrack_Click
        If KeyCode = 77 Then Player_Mute    'M mute
        If KeyCode = 83 Then Button_Click (8)    'S shuffle/Random
    Else
        If Shift = 2 And KeyCode = 79 Then frmPLST.ADD_MULTIPLE_FILES (False)    'O open file
        If Shift = 2 And KeyCode = 82 Then Button_Click (23)
        If Shift = 2 And KeyCode = 69 Then Button_Click (24)    'eq
        If Shift = 2 And KeyCode = 76 Then Button_Click (26)    'plst
        If Shift = 2 And KeyCode = 78 Then frmPLST.clist.Clear: frmPLST.picList.cls: CurrentTrack_Index = -1
        If Shift = 4 And KeyCode = 115 Then Unload frmMain

    End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_Load()
    On Error Resume Next
    'If App.PrevInstance = True Then Exit Sub
    frameContainer.Left = 30
    frameContainer.Top = 30
    PlayerIsPlaying = "false"
    activeSlider = -1
    mSnapDistance = 10 * Screen.TwipsPerPixelX
    ' Change the formicon to display Alpha Icons from resource file
    Call SetIcon(Me.hwnd, "APPICON", False)
    FMOD_Initialize 500, 44100, 5, FSOUND_INIT_ENABLESYSTEMCHANNELFX, FSOUND_OUTPUT_DSOUND, FSOUND_MIXER_QUALITY_AUTODETECT, 0
    Set tCurrentID3 = New cMP3
    Set ttDemo = New Tooltip
    Set cImage = New c32bppDIB
    Call cImage.InitializeDIB(125, 140)

    Set cExToolTip = New ExToolTip             '//-- Creates a new instance of theclass.
    cExToolTip.Start (ScrollText(1).picHandle)
    cExToolTip.Start (ScrollText(5).picHandle)

    Spectrum_Enable True
    'Load_Settings_INI True
    Dim strRes As String
    Dim i As Long
    strRes = Read_INI("Equalizer", "Enabled", 0, , True)
    If CBool(strRes) = True Then frmMain.Button_Click (16)

    For i = 0 To 9
        strRes = Read_INI("Equalizer", "EQ_" & i, 0, , True)
        EqTempValues(i) = CInt(strRes)
        frmMain.Eq_SliderCtrl(i).Value = CInt(strRes)    '.sldEQ(i).Value
    Next

    DrawEQ
    blnLoaded = True
    If PlayerIsPlaying = False Then
        slider(0).Value = 0
        slider(3).Value = 0
    End If

hell:
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    PlayerisClosing = True
    If boolDspShow = True And boolDspLoaded Then frmDSP.Visible = False
    If boolVisShow = True And boolVisLoaded Then FrmVisualisation.Hide
    If boolOptionsShow = True And boolOptionsLoaded Then frmOptions.Hide
    If boolTagsShow = True And boolTagsLoaded Then frmTags.Hide
    If boolMediaLibraryShow = True Then frmLibrary.Hide

    frmLibrary.Visible = False
    Save_Settings_INI True

    '/* remove window from subclassing
    Call Unhook(Me.hwnd)
    Call DisableFileDrops
    Call Unload(frmLibrary)
    Stop_Player

    Dim i As Integer
    Dim tTrack As FileTrack

    Kill (App.Path + "\mplayerlist.npl")
    Open App.Path + "\mplayerlist.npl" For Random As #1 Len = 255
    For i = 1 To frmPLST.clist.ItemCount Step 1
        tTrack.trackName = frmPLST.clist.Item(i - 1)
        tTrack.trackPath = frmPLST.clist.exItem(i - 1)
        tTrack.Duration = frmPLST.clist.exTracklength(i - 1)
        Put #1, i, tTrack
    Next
    Close #1

    If boolVisLoaded = True Then Unload FrmVisualisation
    If boolTagsLoaded = True Then Unload frmTags
    If boolOptionsLoaded = True Then Unload frmOptions
    If boolDspLoaded = True Then Unload frmDSP
    If boolSearchShow Then Unload frmSearch
    If boolMediaLibraryShow Then Unload frmLibrary

    Unload frmPLST
    If boolDspShow = True Then Unload frmDSP
    Me.Hide
    If Dir(DirectoriOWindowS & "MusicMp3.bmp") <> "" Then
        Kill DirectoriOWindowS & "MusicMp3.bmp"
    End If

    'Set frmOptions = Nothing
    Unload frmPopUp
    Set frmMain = Nothing
    Set cExToolTip = Nothing                      '//-- Release all the system and memory resources associated with Class.

    Set ttDemo = Nothing
    Set tCurrentID3 = Nothing
    FMOD_Terminate

    End

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picMiniMode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo err:
    If Button = vbLeftButton Then FormDrag_Down X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
    If frmPLST.Left >= frmMain.Left + frmMain.Width - iBorder - 250 And frmPLST.Left <= frmMain.Left + frmMain.Width + 80 Then
        frmPLST.Left = frmMain.Left + frmMain.Width - iBorder / 2 - 30
        attach1 = True
    ElseIf frmMain.Left >= frmPLST.Left + frmPLST.Width - 200 And frmMain.Left <= frmPLST.Left + frmPLST.Width + 30 Then
        frmPLST.Left = frmMain.Left - frmPLST.Width + iBorder + 15
        attach1 = True
    ElseIf frmPLST.Top >= frmMain.Top + frmMain.Height - iBorder - 200 And frmPLST.Top <= frmMain.Top + frmMain.Height + 200 Then
        frmPLST.Top = frmMain.Top + frmMain.Height - iBorder - 30
        attach1 = True
    ElseIf frmMain.Top >= frmPLST.Top + frmPLST.Height - iTitleHeight - 180 And frmMain.Top <= frmPLST.Top + frmPLST.Height + 100 Then
        frmPLST.Top = frmMain.Top - frmPLST.Height + iTitleHeight + 15
        attach1 = True
    Else
        attach1 = False
    End If

    If FrmVisualisation.Left >= frmMain.Left + frmMain.Width - iBorder - 250 And FrmVisualisation.Left <= frmMain.Left + frmMain.Width + 80 Then
        FrmVisualisation.Left = frmMain.Left + frmMain.Width - iBorder / 2 - 30
        attach2 = True
    ElseIf frmMain.Left >= FrmVisualisation.Left + FrmVisualisation.Width - 200 And frmMain.Left <= FrmVisualisation.Left + FrmVisualisation.Width + 30 Then
        FrmVisualisation.Left = frmMain.Left - FrmVisualisation.Width + iBorder + 15
        attach2 = True
    ElseIf FrmVisualisation.Top >= frmMain.Top + frmMain.Height - iBorder - 200 And FrmVisualisation.Top <= frmMain.Top + frmMain.Height + 200 Then
        FrmVisualisation.Top = frmMain.Top + frmMain.Height - iBorder - 30
        attach2 = True
    ElseIf frmMain.Top >= FrmVisualisation.Top + FrmVisualisation.Height - iTitleHeight - 180 And frmMain.Top <= FrmVisualisation.Top + FrmVisualisation.Height + 100 Then
        FrmVisualisation.Top = frmMain.Top - FrmVisualisation.Height + iTitleHeight + 15
        attach2 = True
    Else
        attach2 = False
    End If

    Exit Sub
err:
    bolDragMini = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picMiniMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    FormDrag_Move X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
End Sub

'-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picMiniMode_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bolDragMini = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub picNormalMode_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    attach1 = False
    attach2 = False

    If Button = vbLeftButton Then FormDrag_Down X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    If Button = vbRightButton Then Me.PopupMenu frmPopUp.mnuMenuPrincipal
    If frmPLST.Left >= frmMain.Left + frmMain.Width - iBorder - 250 And frmPLST.Left <= frmMain.Left + frmMain.Width + 80 Then
        frmPLST.Left = frmMain.Left + frmMain.Width - iBorder / 2 - 30
        attach1 = True
    ElseIf frmMain.Left >= frmPLST.Left + frmPLST.Width - 200 And frmMain.Left <= frmPLST.Left + frmPLST.Width + 30 Then
        frmPLST.Left = frmMain.Left - frmPLST.Width + iBorder + 15
        attach1 = True
    ElseIf frmPLST.Top >= frmMain.Top + frmMain.Height - iBorder - 200 And frmPLST.Top <= frmMain.Top + frmMain.Height + 200 Then
        frmPLST.Top = frmMain.Top + frmMain.Height - iBorder - 30
        attach1 = True
    ElseIf frmMain.Top >= frmPLST.Top + frmPLST.Height - iTitleHeight - 180 And frmMain.Top <= frmPLST.Top + frmPLST.Height + 100 Then
        frmPLST.Top = frmMain.Top - frmPLST.Height + iTitleHeight + 15
        attach1 = True
        'Else
        'attach1 = False
    End If
    If boolVisShow = False Then Exit Sub
    If FrmVisualisation.Left >= frmMain.Left + frmMain.Width - iBorder - 250 And FrmVisualisation.Left <= frmMain.Left + frmMain.Width + 80 Then
        FrmVisualisation.Left = frmMain.Left + frmMain.Width - iBorder / 2 - 30
        attach2 = True
    ElseIf frmMain.Left >= FrmVisualisation.Left + FrmVisualisation.Width - 200 And frmMain.Left <= FrmVisualisation.Left + FrmVisualisation.Width + 30 Then
        FrmVisualisation.Left = frmMain.Left - FrmVisualisation.Width + iBorder + 15
        attach2 = True
    ElseIf FrmVisualisation.Top >= frmMain.Top + frmMain.Height - iBorder - 200 And FrmVisualisation.Top <= frmMain.Top + frmMain.Height + 200 Then
        FrmVisualisation.Top = frmMain.Top + frmMain.Height - iBorder - 30
        attach2 = True
    ElseIf frmMain.Top >= FrmVisualisation.Top + FrmVisualisation.Height - iTitleHeight - 180 And frmMain.Top <= FrmVisualisation.Top + FrmVisualisation.Height + 100 Then
        FrmVisualisation.Top = frmMain.Top - FrmVisualisation.Height + iTitleHeight + 15
        attach2 = True
        'Else
        ' attach2 = False
    End If

    If boolVisShow = False Then Exit Sub
    If FrmVisualisation.Left >= frmPLST.Left + frmPLST.Width - iBorder - 250 And FrmVisualisation.Left <= frmPLST.Left + frmPLST.Width + 80 Then
        FrmVisualisation.Left = frmPLST.Left + frmPLST.Width - iBorder / 2 + 30
        attach2 = True
    ElseIf frmPLST.Left >= FrmVisualisation.Left + FrmVisualisation.Width - 200 And frmPLST.Left <= FrmVisualisation.Left + FrmVisualisation.Width + 30 Then
        FrmVisualisation.Left = frmPLST.Left - FrmVisualisation.Width + iBorder + 15
        attach2 = True
    ElseIf FrmVisualisation.Top >= frmPLST.Top + frmPLST.Height - iBorder - 200 And FrmVisualisation.Top <= frmPLST.Top + frmPLST.Height + 200 Then
        FrmVisualisation.Top = frmPLST.Top + frmPLST.Height - iBorder + 80    '- 30
        attach2 = True
    ElseIf frmPLST.Top >= FrmVisualisation.Top + FrmVisualisation.Height - iTitleHeight - 180 And frmPLST.Top <= FrmVisualisation.Top + FrmVisualisation.Height + 100 Then
        FrmVisualisation.Top = frmPLST.Top - FrmVisualisation.Height + iTitleHeight + 15    '+ 15
        attach2 = True
    End If


    If attach1 And attach2 Then Exit Sub

    If attach2 = True And boolVisShow = True Then
        If frmPLST.Left >= FrmVisualisation.Left + FrmVisualisation.Width And frmPLST.Left <= FrmVisualisation.Left + FrmVisualisation.Width Then
            frmPLST.Left = FrmVisualisation.Left + FrmVisualisation.Width
            attach1 = True
        ElseIf FrmVisualisation.Left >= frmPLST.Left + frmPLST.Width And FrmVisualisation.Left <= frmPLST.Left + frmPLST.Width Then
            frmPLST.Left = FrmVisualisation.Left - frmPLST.Width
            attach1 = True
        ElseIf frmPLST.Top >= FrmVisualisation.Top + FrmVisualisation.Height And frmPLST.Top <= FrmVisualisation.Top + FrmVisualisation.Height + 200 Then
            frmPLST.Top = FrmVisualisation.Top + FrmVisualisation.Height
            attach1 = True
        ElseIf FrmVisualisation.Top >= frmPLST.Top + frmPLST.Height And FrmVisualisation.Top <= frmPLST.Top + frmPLST.Height + 200 Then
            frmPLST.Top = FrmVisualisation.Top - frmPLST.Height
            attach1 = True
        End If
    End If

    'If attach1 And attach2 Then Exit Sub

    'If attach1 = True And boolVisShow = True Then
    'If FrmVisualisation.Left >= frmPLST.Left + frmPLST.Width And FrmVisualisation.Left <= frmPLST.Left + frmPLST.Width Then
    'FrmVisualisation.Left = frmPLST.Left + frmPLST.Width '- iBorder / 2 - 30
    'attach2 = True
    'lseIf frmPLST.Left >= FrmVisualisation.Left + FrmVisualisation.Width And frmPLST.Left <= FrmVisualisation.Left + FrmVisualisation.Width Then
    'FrmVisualisation.Left = frmPLST.Left - FrmVisualisation.Width '+ iBorder + 15
    'attach2 = True
    'ElseIf FrmVisualisation.Top >= frmPLST.Top + frmPLST.Height And FrmVisualisation.Top <= frmPLST.Top + frmPLST.Height Then
    'FrmVisualisation.Top = frmPLST.Top + frmPLST.Height '- iBorder - 30
    'attach2 = True
    'ElseIf frmPLST.Top >= FrmVisualisation.Top + FrmVisualisation.Height And frmPLST.Top <= FrmVisualisation.Top + FrmVisualisation.Height Then
    'FrmVisualisation.Top = frmPLST.Top - FrmVisualisation.Height '+ iTitleHeight + 15
    'attach2 = True
    'End If

    'End If
    Exit Sub
err:
    bolDragMini = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub ButtonMini_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = vbRightButton Then Exit Sub
    Select Case Index
    Case 0: Previous_Track
    Case 1: Play
    Case 2: Pause_Play
    Case 3: Stop_Player
    Case 4: Next_Track
    Case 5: Me.PopupMenu frmPopUp.mnuMenuPrincipal
    Case 6: Minimize_Me
    Case 7: Change_Mask False, True: frmMain.Image_State_Rep
    Case 8: Unload frmMain
    Case 9: Button_Click (26)
    Case 10: Button_Click (23)
    Case 11:
        ButtonMini(Index).Selected = Not ButtonMini(Index).Selected: bCrossFadeEnabled = ButtonMini(Index).Selected:
        If ButtonMini(11).Selected Then
            Show_Message "CROSSFADE ON"
        Else
            Show_Message "CROSSFADE OFF"
        End If
    Case 12: Button_Click (7)    'repeat
    Case 13: Button_Click (8)    'shuffle
    Case 14: Call frmPLST.ADD_MULTIPLE_FILES(True)    'true is set if new list is to be constructed
    End Select
End Sub

Private Sub ScrollText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not MMove Or cExToolTip.Alive = False And (Index = 1 Or Index = 5) Then
        MMove = True
        InitializeToolTip (Index)
    End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Slider_Change(Index As Integer, Value As Long)
    Dim intPercentage As Integer

    On Error Resume Next
    Select Case Index
    Case 1, 4    '// volume bar

        intPercentage = (Value * 100) / 255

        frmPopUp.mnuVolumen.Caption = "Volume" & " [ " & intPercentage & " % ]"
        Call RenameSystemMenu(12, "Volume" & " [ " & intPercentage & " % ]")
        If bMiniMask = False Then
            Show_Message "Volume " & intPercentage & " %"
        Else
            Show_Message "Volume " & intPercentage & " %"
        End If
        VolumeNActuaL = Value
        Stream_SetVolume lCurrentChannel, VolumeNActuaL
    Case 0, 3    '// pos Bar
        If PlayerIsPlaying = "false" Then Exit Sub
        'bSlider = True
        If bSlider = True Then Show_Message Convert_Time_to_string(slider(Index).Value) + "/" + Convert_Time_to_string(slider(Index).Max) + "(" + CStr(CInt((slider(Index).Value / slider(Index).Max) * 100)) + "%)"

    Case 2    '// pos in list rep normal mode
        'ListRep.TopIndex = slider(2).Max - CInt(slider(2).Value)
    Case 5:
        ' Dim k As Double
        Stream_SetBalance lCurrentChannel, Value
        intPercentage = 2 * (Value * 100 / 255 - 50)    '* 100) / 255
        If intPercentage < 0 Then
            Show_Message "Balance:" & -intPercentage & "% " & "Left"
        ElseIf intPercentage > 0 Then
            Show_Message "Balance:" & intPercentage & "% " & "Right"
        Else
            Show_Message "Balance:" & intPercentage & "% " & "Center"
        End If
        'Timer_Wait.Enabled = True
    End Select

End Sub

Private Sub slider_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    'index = Slider_Selected
    'keycode = KeyAscii
    If KeyCode = &H28 Then    ' - volumen keycode=40
        VolumeNActuaL = VolumeNActuaL - 4
        If VolumeNActuaL < 0 Then VolumeNActuaL = 0
        If VolumeNActuaL > 255 Then VolumeNActuaL = 255
        slider(1).Value = VolumeNActuaL
        slider(4).Value = VolumeNActuaL
        Slider_Change 1, VolumeNActuaL
        Slider_Change 4, VolumeNActuaL
    End If
    If KeyCode = &H26 Then    ' + volumen'keycode=38
        VolumeNActuaL = VolumeNActuaL + 4
        If VolumeNActuaL < 0 Then VolumeNActuaL = 0
        If VolumeNActuaL > 255 Then VolumeNActuaL = 255
        slider(1).Value = VolumeNActuaL
        slider(4).Value = VolumeNActuaL
        Slider_Change 1, VolumeNActuaL
        Slider_Change 4, VolumeNActuaL
    End If

    If KeyCode = &H25 Then Five_Seg_Backward  'A Atras 5 seg  keycode=37
    If KeyCode = &H27 Or KeyCode = 100 Then Five_Seg_Forward    'D Adelante 5 seg keycode=39

End Sub

Private Sub Slider_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
' pos bar sliders
    On Error Resume Next
    If PlayerIsPlaying = "false" Then Exit Sub

    If Index = 0 Or Index = 3 Then bSlider = True: Show_Message Convert_Time_to_string(slider(Index).Value) + "/" + Convert_Time_to_string(slider(Index).Max) + "(" + CStr(CInt((slider(Index).Value / slider(Index).Max) * 100)) + "%)"

End Sub

Private Sub Slider_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Slider_Selected = Index
    If Index <> 3 Then slider(Index).SetFocus
    ' If Index = 3 Then frmMain.ListRep.SetFocus

    If PlayerIsPlaying = "false" Then Exit Sub
    Select Case Index
    Case 0, 3    '// pos bar
    End Select

End Sub

Private Sub Slider_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
    Case 1, 4    '// volume bar
        Timer_Wait.Enabled = True
    Case 0, 3    '// pos bar
        If PlayerIsPlaying = "false" Then
            If bMiniMask = True Then
                slider(0).Value = 0
            Else
                slider(3).Value = 0
            End If
            Exit Sub
        Else
            If bMiniMask = True Then
                If bolTimeAct = False Then
                    ScrollText(4).CaptionText = Convert_Time_to_string(slider(Index).Value)
                Else
                    ScrollText(4).CaptionText = "-" & Convert_Time_to_string(slider(Index).Max - slider(Index).Value)
                End If
            Else
                If bolTimeAct = False Then
                    ScrollText(0).CaptionText = Convert_Time_to_string(slider(Index).Value)
                Else
                    ScrollText(0).CaptionText = "-" & Convert_Time_to_string(slider(Index).Max - slider(Index).Value)
                End If
            End If
        End If
        bSlider = False
        'Show_Message Convert_Time_to_string(slider(Index).Value) + "/" + Convert_Time_to_string(slider(Index).Max) + "(" + CStr(CInt((slider(Index).Value / slider(Index).Max) * 100)) + "%)"
        Stream_SetPosition lCurrentChannel, CLng(slider(Index).Value)
    End Select

End Sub

Private Sub Timer_Crossfade_Timer()
    On Error Resume Next
    'Draw_Spectrum

    Select Case peCrossFadeType
    Case 0    '// Crossfade normal
        lVol = lVol - 5
        If lVol <= 0 Or bCrossFadeEnabled = False Then
            Stream_Stop lChannelOut
            Timer_Crossfade.Enabled = False
            lVol = 0
        End If

        Stream_SetVolume lChannelOut, lVol
        Stream_SetVolume lChannelIn, Abs(VolumeNActuaL - lVol)

    Case 1    '// Fade in
        lVol = lVol - 5
        If lVol <= 0 Or bCrossFadeEnabled = False Then
            If PlayerIsPlaying = "false" Then
                Stream_Stop lCurrentChannel
                Stop_Draw_Spectrum
            End If
            If PlayerIsPlaying = "pause" Then Stream_Pause lCurrentChannel
            Timer_Crossfade.Enabled = False
        End If

        Stream_SetVolume lCurrentChannel, lVol

    Case 2    '// Fade Out
        lVol = lVol + 5
        If lVol >= VolumeNActuaL Or bCrossFadeEnabled = False Then Timer_Crossfade.Enabled = False: lVol = VolumeNActuaL

        Stream_SetVolume lCurrentChannel, lVol

    End Select



End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

Private Sub Timer_Wait_Timer()
    If bMiniMask = True Then
        ScrollText(5).CaptionText = sTextScroll
    Else
        ScrollText(1).CaptionText = sTextScroll
    End If
    Timer_Wait.Enabled = False
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Timer_Intro_Timer()
    TiempoIntro = TiempoIntro + 1
    If TiempoIntro = 10 Then
        If PlayerLoop = True Then
            Play
        Else
            Next_Track
        End If
        TiempoIntro = 0
    End If
End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Private Sub Timer_Player_Timer()
    Dim iTimeCross As Integer
    If bLoading = True Or bLoadingSkin = True Then Exit Sub
    '//si esta reproduciendo
    'On Error Resume Next
    If (PlayerIsPlaying = "false") Then Exit Sub
    If (PlayerIsPlaying <> "false" Or (Timer_Crossfade.Enabled = True)) And Me.WindowState <> vbMinimized Then
        'DRW_BARWIDTH = CInt(0.82 * PicSpectrum.ScaleWidth / FFT_BANDS)  ' - 2 * FFT_BANDS
        'If DRW_BARWIDTH < 0 Then DRW_BARWIDTH = 1
        If bMiniMask = False Then
            Draw_Spectrum Me.PicSpectrum
        Else
            Draw_Spectrum Me.PicSpectrumMini
        End If
        'If boolVisShow Then
        'oPlugIn.drawVis FrmVisualisation.ScopeBuff.hdc, InData, FrmVisualisation.ScopeBuff.ScaleHeight, ScopeBuff.ScaleWidth
        'FrmVisualisation.PicSpectrum.Picture = FrmVisualisation.ScopeBuff.Image
    End If

    'On Error Resume Next
    If PlayerIsPlaying = "true" Then

        '// si se esta arrastrando el slider rep
        If bSlider = True Then Exit Sub

        If Stream_GetPosition(lCurrentChannel) > 10 And iCrossfadeTrack <> 0 Then iTimeCross = 5

        '// duracion de la rola
        If bCrossFadeEnabled And Stream_GetDuration(lCurrentChannel) - Stream_GetPosition(lCurrentChannel) <= iTimeCross Or (Not bCrossFadeEnabled And Stream_GetDuration(lCurrentChannel) - Stream_GetPosition(lCurrentChannel) <= 1) Then

            '// si esta seleccionada el check para el loop
            ' If PlayerLoop = True Then Play: Exit Sub
            If PlayerLoop = True And frmPLST.clist.ItemCount <> 0 And CurrentTrack_Index = frmPLST.clist.ItemCount Then CurrentTrack_Index = 0
            Next_Track
            Exit Sub
        End If

        '// Si esta la minimaskara
        If bMiniMask = True Then
            If bolTimeAct = False Then
                ScrollText(4).CaptionText = Convert_Time_to_string(Stream_GetPosition(lCurrentChannel))
            Else
                ScrollText(4).CaptionText = "-" & Convert_Time_to_string(Stream_GetDuration(lCurrentChannel) - Stream_GetPosition(lCurrentChannel))
            End If
            slider(3).Value = CInt(Stream_GetPosition(lCurrentChannel))
        Else
            If bolTimeAct = False Then
                ScrollText(0).CaptionText = Convert_Time_to_string(Stream_GetPosition(lCurrentChannel))
            Else
                ScrollText(0).CaptionText = "-" & Convert_Time_to_string(Stream_GetDuration(lCurrentChannel) - Stream_GetPosition(lCurrentChannel))
            End If
            slider(0).Value = CInt(Stream_GetPosition(lCurrentChannel))
        End If


        'If tCurrentID3.Lyrics3Tag = True And bolLyricsShow = True Then
    End If


End Sub

Sub FormDrag_Move(X As Single, Y As Single)
    On Error Resume Next
    Dim DiffX As Long, DiffY As Long
    Dim NewX As Long, NewY As Long
    Dim ToLeftDistance As Long
    Dim ToRightDistance As Long
    Dim ToTopDistance As Long
    Dim ToBottomDistance As Long

    If bolDragMini = True Then
        DiffX = X - StartDragX
        DiffY = Y - StartDragY

        If DiffX = 0 And DiffY = 0 Then Exit Sub
        NewX = Me.Left + DiffX
        NewY = Me.Top + DiffY


        ToRightDistance = rWorkArea.Right - (NewX + Me.Width)
        ToLeftDistance = NewX - rWorkArea.Left
        ToBottomDistance = rWorkArea.Bottom - (NewY + Me.Height)
        ToTopDistance = NewY - rWorkArea.Top + iTitleHeight / 2

        If Not mAttachedToBottom Then
            If Abs(ToBottomDistance) <= mSnapDistance Then
                NewY = rWorkArea.Bottom - Me.Height + iBorder * 3 / 2#
                mAttachedToBottom = True
            End If
        Else

            If Abs(ToBottomDistance) > mSnapDistance Then
                mAttachedToBottom = False
            Else
                NewY = Me.Top
            End If
        End If

        If Not mAttachedToTop Then
            If Abs(ToTopDistance) <= mSnapDistance Then
                NewY = rWorkArea.Top - iTitleHeight    '+ iBorder / 2
                mAttachedToTop = True
            End If
        Else
            If Abs(ToTopDistance) > mSnapDistance Then
                mAttachedToTop = False
            Else
                NewY = Me.Top
            End If
        End If

        If Not mAttachedToRight Then
            If Abs(ToRightDistance) <= mSnapDistance Then
                NewX = rWorkArea.Right - Me.Width + iBorder / 2
                mAttachedToRight = True
            End If
        Else
            If Abs(ToRightDistance) > mSnapDistance Then
                mAttachedToRight = False
            Else
                NewX = Me.Left
            End If
        End If

        If Not mAttachedToLeft Then
            If Abs(ToLeftDistance) <= mSnapDistance Then
                NewX = rWorkArea.Left - iBorder
                mAttachedToLeft = True
            End If
        Else
            If Abs(ToLeftDistance) > mSnapDistance Then
                mAttachedToLeft = False
            Else
                NewX = Me.Left
            End If
        End If

        If attach1 = True Then
            DiffX = NewX - frmMain.Left
            DiffY = NewY - frmMain.Top
            frmPLST.Move frmPLST.Left + DiffX, frmPLST.Top + DiffY
        End If
        If attach2 = True Then
            DiffX = NewX - frmMain.Left
            DiffY = NewY - frmMain.Top
            FrmVisualisation.Move FrmVisualisation.Left + DiffX, FrmVisualisation.Top + DiffY
        End If

        Me.Move NewX, NewY
    End If

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub FormDrag_Down(X As Single, Y As Single)
    On Error Resume Next

    'SystemGetWorkArea SPI_GETWORKAREA, 0&, rWorkArea, 0&

    rWorkArea.Top = rWorkArea.Top * Screen.TwipsPerPixelY
    rWorkArea.Left = rWorkArea.Left * Screen.TwipsPerPixelX
    rWorkArea.Bottom = rWorkArea.Bottom * Screen.TwipsPerPixelY
    rWorkArea.Right = rWorkArea.Right * Screen.TwipsPerPixelX

    bolDragMini = True
    StartDragX = X
    StartDragY = Y

End Sub

Sub Five_Seg_Forward()
    On Error Resume Next
    Dim CurPos As Long
    If PlayerIsPlaying = "false" Then Exit Sub
    If iPosScrollChange <= 0 Then iPosScrollChange = 5
    CurPos = Stream_GetPosition(lCurrentChannel)
    CurPos = CurPos + iPosScrollChange
    If CurPos > Stream_GetDuration(lCurrentChannel) Then CurPos = Stream_GetDuration(lCurrentChannel)
    Stream_SetPosition lCurrentChannel, CurPos
    Show_Message Convert_Time_to_string(CurPos) + "/" + Convert_Time_to_string(CurPos) + "(" + CStr(CInt(((CurPos) / Stream_GetDuration(lCurrentChannel)) * 100)) + "%)"

End Sub

Sub Minimize_Me()
    If boolOptionsShow = True Then frmOptions.Hide
    If boolTagsShow = True Then frmTags.Hide
    If boolDspShow = True Then frmDSP.Hide
    If boolMediaLibraryShow = True Then frmLibrary.Hide

    If OpcionesMusic.SysTray = False And OpcionesMusic.TaskBar = False Then
        OpcionesMusic.TaskBar = True
    End If

    If OpcionesMusic.TaskBar = True Then frmPopUp.WindowState = vbMinimized: Me.WindowState = vbMinimized
    If OpcionesMusic.SysTray = True Then Me.Hide

End Sub

Public Sub Mouse_Wheel_Moving(Mouse_Wheel_down As Boolean)
    On Error Resume Next
    Dim Index As Integer
    'just to chech the control on slider
    If iVolScrollChange <= 0 Then iVolScrollChange = 4
    Index = Slider_Selected
    If Mouse_Wheel_down Then    ' + volume
        If Index = 1 Or Index = 4 Then
            VolumeNActuaL = VolumeNActuaL + iVolScrollChange
            If VolumeNActuaL < 0 Then VolumeNActuaL = 0
            If VolumeNActuaL > 255 Then VolumeNActuaL = 255
            slider(Index).Value = VolumeNActuaL
            Slider_Change Index, VolumeNActuaL
        ElseIf Index = 0 Or Index = 3 Then
            Five_Seg_Forward
        End If

    Else    ' - volume
        If Index = 1 Or Index = 4 Then
            VolumeNActuaL = VolumeNActuaL - iVolScrollChange
            If VolumeNActuaL < 0 Then VolumeNActuaL = 0
            If VolumeNActuaL > 255 Then VolumeNActuaL = 255
            slider(Index).Value = VolumeNActuaL
            Slider_Change Index, VolumeNActuaL
        ElseIf Index = 0 Or Index = 3 Then
            Five_Seg_Backward
        End If
    End If


    On Error Resume Next
    Me.SetFocus
End Sub

Public Sub EQ_in_out()
    Dim newRegion

    If frmMain.Button(24).Selected = True Then

        If eq_position = 0 Then          'HORIZONTAL
            If attach1 = True And frmPLST.Left >= frmMain.Left + frmMain.Width - 3 * iBorder Then
                frmPLST.Left = frmMain.Left + Width_With_Eq + 15
            End If
            Me.Width = Width_With_Eq + iBorder
        ElseIf eq_position = 1 Then      'VERTICAL

            If attach1 = True And frmPLST.Top >= frmMain.Top + frmMain.Height - iTitleHeight - 2 * iBorder Then
                frmPLST.Top = frmMain.Top + Width_With_Eq + iTitleHeight - 30
            End If
            Me.Height = Width_With_Eq + iTitleHeight + iBorder
        End If
        newRegion = ExtCreateRegionByte(ByVal 0&, EdgeRegions(0).DataLength, EdgeRegions(0).RegionData(0))
        SetWindowRgn frmMain.hwnd, newRegion, True
    Else
        If eq_position = 0 Then          'HORIZONTAL
            If attach1 = True And frmPLST.Left >= frmMain.Left + frmMain.Width - 2 * iBorder Then
                frmPLST.Left = frmMain.Left + Width_Without_Eq + iBorder - iBorder / 2 - 30
            End If
            Me.Width = Width_Without_Eq + iBorder
        ElseIf eq_position = 1 Then      'VERTICAL
            If attach1 = True And frmPLST.Top >= frmMain.Top + frmMain.Height - iTitleHeight - 2 * iBorder Then
                frmPLST.Top = frmMain.Top + Width_Without_Eq + iTitleHeight - 30
            End If
            Me.Height = Width_Without_Eq + iTitleHeight + iBorder
        End If
        newRegion = ExtCreateRegionByte(ByVal 0&, EdgeRegions(2).DataLength, EdgeRegions(2).RegionData(0))
        SetWindowRgn frmMain.hwnd, newRegion, True

    End If
    DeleteObject newRegion
    EqBackpic.Refresh

    'If bLoading = False Or bLoadingSkin = False Then frmMain.Show
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RANDOM_track
' Author    : Mahesh Kurmi
' Date      : 3/3/2011
' Purpose   : Shuffle tracks
' HOW       : It randomises plsylist snd stores in an array
' Returns   :
' Parameters:
'---------------------------------------------------------------------------------------
'
Public Sub RANDOM_track()
    Dim j, i
    ReDim Random_Order_track(frmPLST.clist.ItemCount - 1)

    Randomize

    If PlayerIsPlaying = "false" Then
        Random_Order_track(0) = Int(frmPLST.clist.ItemCount * Rnd)
    Else
        Random_Order_track(0) = CurrentTrack_Index
        If Random_Order_track(0) = -1 Then Random_Order_track(0) = Int(frmPLST.clist.ItemCount * Rnd)
    End If

    For j = 1 To frmPLST.clist.ItemCount - 1
        DoEvents
        Randomize
        Random_Order_track(j) = Int(frmPLST.clist.ItemCount * Rnd)
        For i = 0 To j - 1
            If Random_Order_track(j) = Random_Order_track(i) Then
                j = j - 1
                If j < 1 Then j = 1
                Exit For
            End If
        Next i
    Next j
End Sub

Public Sub EnableCrossfade(bEnable As Boolean)

    frmPopUp.mnuCrossfade.Checked = bEnable
    If bEnable = True Then
        PicCrossfade.PaintPicture PicConfigInfo, 0, 0, PicCrossfade.Width, PicCrossfade.Height, PicShuffle.Width + PicRepeat.Width, PicCrossfade.Height, PicCrossfade.Width, PicCrossfade.Height
        PicCrossfade.Picture = PicCrossfade.Image
        Show_Message ("Crossfade:ON")
    Else
        PicCrossfade.PaintPicture PicConfigInfo, 0, 0, PicCrossfade.Width, PicCrossfade.Height, PicShuffle.Width + PicRepeat.Width, 0, PicCrossfade.Width, PicCrossfade.Height
        PicCrossfade.Picture = PicCrossfade.Image
        Show_Message ("Crossfade:OFF")
    End If
End Sub

Public Sub Setpreset(iJump As Integer)
    Dim sTemp As String
    Dim i As Integer
    iPresetIndex = iPresetIndex + iJump
    If iPresetIndex >= ListEq.ListCount Then iPresetIndex = 0
    If iPresetIndex < 0 Then iPresetIndex = ListEq.ListCount
    sTemp = Right(ListEq.list(iPresetIndex), Len(ListEq.list(iPresetIndex)) - InStr(1, ListEq.list(iPresetIndex), ".", vbBinaryCompare))
    Dim Arr As Variant
    Arr = Split(sTemp, ",")
    For i = 0 To 9
        Eq_SliderCtrl(i).Value = CInt(Arr(i))
    Next
    For i = 0 To 9 Step 1
        FX_SetEQ CLng(i), CLng(Eq_SliderCtrl(i).Value)
    Next

    sTemp = Left(ListEq.list(iPresetIndex), InStr(1, ListEq.list(iPresetIndex), ".", vbBinaryCompare) - 1)
    Show_Message "Preset:" + sTemp
    Button(18).ToolTipText = sTemp

End Sub

Private Sub InitializeToolTip(Index As Integer)
'index :(value=1 normalwindow Textscoll
'index :(value=5 minimask textscroll
    Dim tooltext As String
    tooltext = " Track :  " & Format$(str(CurrentTrack_Index + 1), "00") + "/" + Format$(str(frmPLST.clist.ItemCount), "00") & vbCrLf & _
               " Title    :  " & tCurrentID3.Title & vbCrLf & _
               " Length:  " & tCurrentID3.MPEG_DurationTime & vbCrLf & _
               " Artist  :  " & tCurrentID3.Artist & vbCrLf & _
               " Album :  " & tCurrentID3.Album & vbCrLf & _
               " Year   :  " & tCurrentID3.Year & vbCrLf & _
               " Genre :  " & tCurrentID3.Genre & vbCrLf & _
               " Path    :  " & sFileMainPlaying
    cExToolTip.ToolTipStyle = TTBalloon
    cExToolTip.Shadow = True
    cExToolTip.BackColor = frmPLST.picList.BackColor
    cExToolTip.TextColor = NormalText_Forecolor
    cExToolTip.DelayTime = 30
    cExToolTip.IconSize = TTIcon125
    'Load image of album art if available
    If frmMain.cImage.ImageType = imgError Then cExToolTip.IconSize = TTIcon1    'set very samll size to image not to be duspalyed
    'cExToolTip.Font = frmPLST.Font
    cExToolTip.ShowToolTip ScrollText(Index).picHandle, "Mahesh Mp3Player", tooltext, frmMain.PicConfigInfo.Picture, 100
End Sub

