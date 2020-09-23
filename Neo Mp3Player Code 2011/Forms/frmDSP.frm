VERSION 5.00
Begin VB.Form frmDSP 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13275
   ControlBox      =   0   'False
   Icon            =   "frmDSP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   622
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MMPlayerXProject.vkFrame frameBack 
      Height          =   5535
      Index           =   0
      Left            =   300
      TabIndex        =   0
      Top             =   555
      Width           =   8355
      _extentx        =   14737
      _extenty        =   9763
      backcolor1      =   14737632
      font            =   "frmDSP.frx":000C
      showtitle       =   0   'False
      titlecolor1     =   12632256
      bordercolor     =   0
      roundangle      =   0
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3960
         Index           =   1
         Left            =   0
         TabIndex        =   7
         Top             =   1020
         Width           =   8355
         _extentx        =   14737
         _extenty        =   6985
         backcolor1      =   14737632
         backcolor2      =   12632256
         font            =   "frmDSP.frx":0034
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin MMPlayerXProject.vkCommand vkBtnEQ 
            Height          =   405
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Top             =   2040
            Width           =   2415
            _extentx        =   4260
            _extenty        =   714
            caption         =   "Reset EQ"
            font            =   "frmDSP.frx":005C
         End
         Begin VB.TextBox txtEQPreset 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   3780
            TabIndex        =   14
            Top             =   2040
            Width           =   4425
         End
         Begin MMPlayerXProject.vkListBox vklistEQ 
            Height          =   1500
            Left            =   3780
            TabIndex        =   13
            Top             =   2370
            Width           =   4425
            _extentx        =   7805
            _extenty        =   2646
            backcolor       =   14737632
            bordercolor     =   0
            multiselect     =   0   'False
            sorted          =   0
            stylecheckbox   =   -1  'True
            font            =   "frmDSP.frx":0084
            selcolor        =   12632256
            borderselcolor  =   12632256
         End
         Begin MMPlayerXProject.vkCommand vkBtnEQ 
            Height          =   405
            Index           =   1
            Left            =   150
            TabIndex        =   11
            Top             =   2520
            Width           =   2415
            _extentx        =   4260
            _extenty        =   714
            caption         =   "Save EQ"
            font            =   "frmDSP.frx":00AC
         End
         Begin MMPlayerXProject.vkCommand vkBtnEQ 
            Height          =   405
            Index           =   3
            Left            =   150
            TabIndex        =   10
            Top             =   3480
            Width           =   2415
            _extentx        =   4260
            _extenty        =   714
            caption         =   "Reload Default Presets"
            font            =   "frmDSP.frx":00D4
         End
         Begin MMPlayerXProject.vkCommand vkBtnEQ 
            Height          =   405
            Index           =   2
            Left            =   150
            TabIndex        =   9
            Top             =   3000
            Width           =   2415
            _extentx        =   4260
            _extenty        =   714
            caption         =   "Delete EQ"
            font            =   "frmDSP.frx":00FC
         End
         Begin VB.PictureBox PicContainer 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   1845
            Index           =   0
            Left            =   150
            ScaleHeight     =   121
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   534
            TabIndex        =   8
            Top             =   120
            Width           =   8045
            Begin MMPlayerXProject.Eq_SliderCtrl sldEQ 
               Height          =   1380
               Index           =   0
               Left            =   3840
               TabIndex        =   12
               Top             =   150
               Width           =   120
               _extentx        =   2434
               _extenty        =   212
               pictureback     =   "frmDSP.frx":0124
               pictureprogress =   "frmDSP.frx":0CF6
               bardown         =   "frmDSP.frx":18C8
               barover         =   "frmDSP.frx":1BDA
               bar             =   "frmDSP.frx":1EEC
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   285
               Index           =   7
               Left            =   120
               TabIndex        =   16
               Top             =   150
               Width           =   1155
               _extentx        =   2037
               _extenty        =   503
               backstyle       =   0
               caption         =   "Enable EQ"
               font            =   "frmDSP.frx":21FE
            End
            Begin MMPlayerXProject.vkCheck vkCheckEqmode 
               Height          =   285
               Index           =   2
               Left            =   120
               TabIndex        =   17
               Top             =   1380
               Width           =   1035
               _extentx        =   1826
               _extenty        =   503
               backstyle       =   0
               font            =   "frmDSP.frx":2226
            End
            Begin MMPlayerXProject.vkCheck vkCheckEqmode 
               Height          =   285
               Index           =   1
               Left            =   120
               TabIndex        =   18
               Top             =   970
               Width           =   1035
               _extentx        =   1826
               _extenty        =   503
               backstyle       =   0
               caption         =   "Locked"
               font            =   "frmDSP.frx":224E
            End
            Begin MMPlayerXProject.vkCheck vkCheckEqmode 
               Height          =   285
               Index           =   0
               Left            =   120
               TabIndex        =   19
               Top             =   560
               Width           =   2115
               _extentx        =   3731
               _extenty        =   503
               backstyle       =   0
               caption         =   "Allow free movement"
               font            =   "frmDSP.frx":2276
            End
            Begin VB.Label lblEQ 
               AutoSize        =   -1  'True
               Caption         =   "16kHz"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   150
               Index           =   0
               Left            =   3720
               TabIndex        =   20
               Top             =   1590
               Visible         =   0   'False
               Width           =   315
            End
         End
      End
      Begin MMPlayerXProject.vkCommand vkBtn 
         Height          =   405
         Index           =   5
         Left            =   150
         TabIndex        =   22
         Top             =   5040
         Width           =   2415
         _extentx        =   4260
         _extenty        =   714
         caption         =   "Default FX Parameters"
         font            =   "frmDSP.frx":229E
      End
      Begin MMPlayerXProject.vkCommand vkBtn 
         Height          =   405
         Index           =   6
         Left            =   2820
         TabIndex        =   21
         Top             =   5040
         Width           =   2205
         _extentx        =   3889
         _extenty        =   714
         caption         =   "Clear All FX "
         font            =   "frmDSP.frx":22C6
      End
      Begin MMPlayerXProject.vkCommand vkBtn1 
         Height          =   405
         Index           =   2
         Left            =   1800
         TabIndex        =   1
         Top             =   4470
         Width           =   1395
         _extentx        =   2461
         _extenty        =   714
         caption         =   "Apply"
         font            =   "frmDSP.frx":22EE
      End
      Begin MMPlayerXProject.vkCommand vkBtn 
         Height          =   405
         Index           =   4
         Left            =   6645
         TabIndex        =   5
         Top             =   630
         Width           =   1695
         _extentx        =   2990
         _extenty        =   714
         backcolor2      =   12632256
         caption         =   "ECHO"
         font            =   "frmDSP.frx":2316
         breakcorner     =   0   'False
         drawfocus       =   0   'False
         drawmouseinrect =   0   'False
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand vkBtn 
         Height          =   405
         Index           =   3
         Left            =   4980
         TabIndex        =   4
         Top             =   630
         Width           =   1680
         _extentx        =   2963
         _extenty        =   714
         backcolor2      =   12632256
         caption         =   "DISTORTION"
         font            =   "frmDSP.frx":233A
         breakcorner     =   0   'False
         drawfocus       =   0   'False
         drawmouseinrect =   0   'False
         disabledbackcolor=   14737632
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand vkBtn 
         Height          =   405
         Index           =   2
         Left            =   3330
         TabIndex        =   2
         Top             =   630
         Width           =   1680
         _extentx        =   2963
         _extenty        =   714
         backcolor2      =   12632256
         caption         =   "WAVE REVERB"
         font            =   "frmDSP.frx":235E
         breakcorner     =   0   'False
         drawfocus       =   0   'False
         drawmouseinrect =   0   'False
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand vkBtn 
         Height          =   405
         Index           =   1
         Left            =   1665
         TabIndex        =   3
         Top             =   630
         Width           =   1680
         _extentx        =   2963
         _extenty        =   714
         backcolor2      =   12632256
         caption         =   "I3D2 REVERB"
         font            =   "frmDSP.frx":2382
         breakcorner     =   0   'False
         drawfocus       =   0   'False
         drawmouseinrect =   0   'False
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand vkBtn 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   6
         Top             =   630
         Width           =   1680
         _extentx        =   2963
         _extenty        =   714
         backcolor2      =   12632256
         caption         =   "EQUALIZERS"
         font            =   "frmDSP.frx":23A6
         breakcorner     =   0   'False
         drawfocus       =   0   'False
         drawmouseinrect =   0   'False
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3960
         Index           =   5
         Left            =   0
         TabIndex        =   56
         Top             =   1020
         Width           =   8355
         _extentx        =   14737
         _extenty        =   6985
         backcolor1      =   14737632
         backcolor2      =   12632256
         caption         =   "Clear All FX"
         font            =   "frmDSP.frx":23CA
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin VB.PictureBox PicContainer 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3675
            Index           =   3
            Left            =   150
            ScaleHeight     =   243
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   533
            TabIndex        =   57
            Top             =   120
            Width           =   8025
            Begin MMPlayerXProject.Eq_SliderCtrl sldEcho 
               Height          =   135
               Index           =   0
               Left            =   1530
               TabIndex        =   58
               Top             =   810
               Width           =   1125
               _extentx        =   238
               _extenty        =   1984
               pictureback     =   "frmDSP.frx":23F2
               pictureprogress =   "frmDSP.frx":2C48
               bardown         =   "frmDSP.frx":349E
               barover         =   "frmDSP.frx":36A0
               bar             =   "frmDSP.frx":38A2
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
               position        =   1
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   285
               Index           =   4
               Left            =   3000
               TabIndex        =   59
               Top             =   210
               Width           =   1545
               _extentx        =   2725
               _extenty        =   503
               backstyle       =   0
               caption         =   "Flanger"
               font            =   "frmDSP.frx":3AA4
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   255
               Index           =   3
               Left            =   270
               TabIndex        =   60
               Top             =   210
               Width           =   1575
               _extentx        =   2778
               _extenty        =   450
               backstyle       =   0
               caption         =   "Echo"
               font            =   "frmDSP.frx":3ACC
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   285
               Index           =   5
               Left            =   5520
               TabIndex        =   72
               Top             =   210
               Width           =   1545
               _extentx        =   2725
               _extenty        =   503
               backstyle       =   0
               caption         =   "Gargle"
               font            =   "frmDSP.frx":3AF4
            End
            Begin MMPlayerXProject.Eq_SliderCtrl sldFlan 
               Height          =   135
               Index           =   0
               Left            =   4200
               TabIndex        =   92
               Top             =   780
               Width           =   1125
               _extentx        =   1984
               _extenty        =   238
               pictureback     =   "frmDSP.frx":3B1C
               pictureprogress =   "frmDSP.frx":4372
               bardown         =   "frmDSP.frx":4BC8
               barover         =   "frmDSP.frx":4DCA
               bar             =   "frmDSP.frx":4FCC
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
               position        =   1
            End
            Begin MMPlayerXProject.Eq_SliderCtrl sldGarg 
               Height          =   135
               Index           =   0
               Left            =   6750
               TabIndex        =   93
               Top             =   780
               Width           =   1125
               _extentx        =   238
               _extenty        =   1984
               pictureback     =   "frmDSP.frx":51CE
               pictureprogress =   "frmDSP.frx":5A24
               bardown         =   "frmDSP.frx":627A
               barover         =   "frmDSP.frx":647C
               bar             =   "frmDSP.frx":667E
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
               position        =   1
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Waveshape:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   47
               Left            =   5490
               TabIndex        =   95
               Top             =   1140
               Width           =   1065
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hz:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   46
               Left            =   5490
               TabIndex        =   94
               Top             =   720
               Width           =   285
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phase:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   45
               Left            =   3000
               TabIndex        =   73
               Top             =   3330
               Width           =   585
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delay:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   44
               Left            =   3000
               TabIndex        =   71
               Top             =   2895
               Width           =   570
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Waveform:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   43
               Left            =   3000
               TabIndex        =   70
               Top             =   2470
               Width           =   960
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frequency:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   42
               Left            =   3000
               TabIndex        =   69
               Top             =   2040
               Width           =   960
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Feedback:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   41
               Left            =   3000
               TabIndex        =   68
               Top             =   1610
               Width           =   885
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Depth:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   40
               Left            =   3000
               TabIndex        =   67
               Top             =   1140
               Width           =   585
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Wet dry mix:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   39
               Left            =   3000
               TabIndex        =   66
               Top             =   720
               Width           =   1125
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pan delay:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   38
               Left            =   270
               TabIndex        =   65
               Top             =   2470
               Width           =   915
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Right delay:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   37
               Left            =   270
               TabIndex        =   64
               Top             =   2040
               Width           =   1035
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Left delay:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   36
               Left            =   270
               TabIndex        =   63
               Top             =   1610
               Width           =   915
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Feedback:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   35
               Left            =   270
               TabIndex        =   62
               Top             =   1140
               Width           =   885
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Wet dry mix:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   34
               Left            =   270
               TabIndex        =   61
               Top             =   720
               Width           =   1125
            End
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3960
         Index           =   4
         Left            =   0
         TabIndex        =   40
         Top             =   1020
         Width           =   8355
         _extentx        =   14737
         _extenty        =   6985
         backcolor1      =   14737632
         backcolor2      =   12632256
         caption         =   "Clear All FX"
         font            =   "frmDSP.frx":6880
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin VB.PictureBox PicContainer 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3675
            Index           =   1
            Left            =   150
            ScaleHeight     =   243
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   533
            TabIndex        =   41
            Top             =   120
            Width           =   8025
            Begin MMPlayerXProject.Eq_SliderCtrl sldComp 
               Height          =   135
               Index           =   0
               Left            =   1860
               TabIndex        =   42
               Top             =   780
               Width           =   1125
               _extentx        =   1984
               _extenty        =   238
               pictureback     =   "frmDSP.frx":68A8
               pictureprogress =   "frmDSP.frx":70FE
               bardown         =   "frmDSP.frx":7954
               barover         =   "frmDSP.frx":7B56
               bar             =   "frmDSP.frx":7D58
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
               position        =   1
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   285
               Index           =   2
               Left            =   4200
               TabIndex        =   43
               Top             =   180
               Width           =   1545
               _extentx        =   2725
               _extenty        =   503
               backstyle       =   0
               caption         =   "Distortion"
               font            =   "frmDSP.frx":7F5A
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   255
               Index           =   1
               Left            =   330
               TabIndex        =   44
               Top             =   180
               Width           =   1575
               _extentx        =   1931
               _extenty        =   450
               backstyle       =   0
               caption         =   "Compression"
               font            =   "frmDSP.frx":7F82
            End
            Begin MMPlayerXProject.Eq_SliderCtrl sldDis 
               Height          =   135
               Index           =   0
               Left            =   5880
               TabIndex        =   90
               Top             =   810
               Width           =   1125
               _extentx        =   238
               _extenty        =   1984
               pictureback     =   "frmDSP.frx":7FAA
               pictureprogress =   "frmDSP.frx":8800
               bardown         =   "frmDSP.frx":9056
               barover         =   "frmDSP.frx":9258
               bar             =   "frmDSP.frx":945A
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
               position        =   1
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lowpass cutoff:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   33
               Left            =   4170
               TabIndex        =   55
               Top             =   2580
               Width           =   1335
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Eq Bandwidth:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   32
               Left            =   4170
               TabIndex        =   54
               Top             =   2115
               Width           =   1230
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Eq center freq:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   31
               Left            =   4170
               TabIndex        =   53
               Top             =   1665
               Width           =   1290
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Edge:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   30
               Left            =   4170
               TabIndex        =   52
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gain:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   29
               Left            =   4170
               TabIndex        =   51
               Top             =   780
               Width           =   465
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pre-delay:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   28
               Left            =   330
               TabIndex        =   50
               Top             =   3030
               Width           =   900
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ratio:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   27
               Left            =   330
               TabIndex        =   49
               Top             =   2580
               Width           =   510
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Threshold:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   26
               Left            =   330
               TabIndex        =   48
               Top             =   2115
               Width           =   915
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Release:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   25
               Left            =   330
               TabIndex        =   47
               Top             =   1665
               Width           =   750
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Attack:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   24
               Left            =   330
               TabIndex        =   46
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gain"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   23
               Left            =   330
               TabIndex        =   45
               Top             =   750
               Width           =   390
            End
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3960
         Index           =   3
         Left            =   0
         TabIndex        =   74
         Top             =   1020
         Width           =   8355
         _extentx        =   14737
         _extenty        =   6985
         backcolor1      =   14737632
         backcolor2      =   12632256
         caption         =   "Clear All FX"
         font            =   "frmDSP.frx":965C
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin VB.PictureBox PicContainer 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3675
            Index           =   2
            Left            =   150
            ScaleHeight     =   243
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   533
            TabIndex        =   75
            Top             =   120
            Width           =   8025
            Begin MMPlayerXProject.Eq_SliderCtrl sldChorus 
               Height          =   135
               Index           =   0
               Left            =   1890
               TabIndex        =   76
               Top             =   780
               Width           =   1125
               _extentx        =   238
               _extenty        =   1984
               pictureback     =   "frmDSP.frx":9684
               pictureprogress =   "frmDSP.frx":9EDA
               bardown         =   "frmDSP.frx":A730
               barover         =   "frmDSP.frx":A932
               bar             =   "frmDSP.frx":AB34
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
               position        =   1
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   285
               Index           =   8
               Left            =   3780
               TabIndex        =   77
               Top             =   180
               Width           =   1515
               _extentx        =   3466
               _extenty        =   503
               backstyle       =   0
               caption         =   "Wave Reverb"
               font            =   "frmDSP.frx":AD36
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   255
               Index           =   0
               Left            =   330
               TabIndex        =   78
               Top             =   180
               Width           =   1095
               _extentx        =   1931
               _extenty        =   450
               backstyle       =   0
               caption         =   "Chorus"
               font            =   "frmDSP.frx":AD5E
            End
            Begin MMPlayerXProject.Eq_SliderCtrl sldWaves 
               Height          =   135
               Index           =   0
               Left            =   6060
               TabIndex        =   91
               Top             =   780
               Width           =   1125
               _extentx        =   1984
               _extenty        =   238
               pictureback     =   "frmDSP.frx":AD86
               pictureprogress =   "frmDSP.frx":B5DC
               bardown         =   "frmDSP.frx":BE32
               barover         =   "frmDSP.frx":C034
               bar             =   "frmDSP.frx":C236
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
               position        =   1
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Wet dry dix:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   12
               Left            =   330
               TabIndex        =   89
               Top             =   750
               Width           =   1065
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Depth:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   13
               Left            =   330
               TabIndex        =   88
               Top             =   1170
               Width           =   585
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Feedback"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   14
               Left            =   330
               TabIndex        =   87
               Top             =   1575
               Width           =   810
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frequency:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   15
               Left            =   330
               TabIndex        =   86
               Top             =   1980
               Width           =   960
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Waveform:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   16
               Left            =   330
               TabIndex        =   85
               Top             =   2385
               Width           =   960
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Delay:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   17
               Left            =   330
               TabIndex        =   84
               Top             =   2805
               Width           =   570
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phase:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   18
               Left            =   330
               TabIndex        =   83
               Top             =   3210
               Width           =   585
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "In gain:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   19
               Left            =   3750
               TabIndex        =   82
               Top             =   750
               Width           =   675
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reverb Mix:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   20
               Left            =   3750
               TabIndex        =   81
               Top             =   1170
               Width           =   1035
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reverb Time:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   21
               Left            =   3750
               TabIndex        =   80
               Top             =   1575
               Width           =   1170
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "High frequency ratio:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   22
               Left            =   3750
               TabIndex        =   79
               Top             =   1980
               Width           =   1815
            End
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3960
         Index           =   2
         Left            =   0
         TabIndex        =   24
         Top             =   1020
         Width           =   8355
         _extentx        =   14737
         _extenty        =   6985
         backcolor1      =   14737632
         backcolor2      =   12632256
         caption         =   "Clear All FX"
         font            =   "frmDSP.frx":C438
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin VB.PictureBox PicContainer 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   3675
            Index           =   4
            Left            =   150
            ScaleHeight     =   243
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   533
            TabIndex        =   25
            Top             =   120
            Width           =   8025
            Begin MMPlayerXProject.Eq_SliderCtrl sldL2 
               Height          =   135
               Index           =   0
               Left            =   1980
               TabIndex        =   26
               Top             =   750
               Width           =   1125
               _extentx        =   238
               _extenty        =   1984
               pictureback     =   "frmDSP.frx":C460
               pictureprogress =   "frmDSP.frx":CCB6
               bardown         =   "frmDSP.frx":D50C
               barover         =   "frmDSP.frx":D70E
               bar             =   "frmDSP.frx":D910
               backcolor       =   -2147483643
               min             =   -10
               max             =   10
               position        =   1
            End
            Begin MMPlayerXProject.vkCheck chkDSP 
               Height          =   285
               Index           =   6
               Left            =   330
               TabIndex        =   27
               Top             =   120
               Width           =   1965
               _extentx        =   3466
               _extenty        =   503
               backstyle       =   0
               caption         =   "I3d Level2 Reverb"
               font            =   "frmDSP.frx":DB12
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Room:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   330
               TabIndex        =   39
               Top             =   720
               Width           =   570
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Room hf:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   330
               TabIndex        =   38
               Top             =   1200
               Width           =   795
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Roll off factor:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   330
               TabIndex        =   37
               Top             =   1710
               Width           =   1230
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Decay time:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   330
               TabIndex        =   36
               Top             =   2175
               Width           =   1050
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Decay hf ratio:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   330
               TabIndex        =   35
               Top             =   2670
               Width           =   1290
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reflections:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   5
               Left            =   330
               TabIndex        =   34
               Top             =   3150
               Width           =   1005
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reflec. delay:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   4530
               TabIndex        =   33
               Top             =   720
               Width           =   1185
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reverb:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   7
               Left            =   4530
               TabIndex        =   32
               Top             =   1200
               Width           =   690
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Reverb delay:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   8
               Left            =   4530
               TabIndex        =   31
               Top             =   1710
               Width           =   1215
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Diffusion:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   9
               Left            =   4530
               TabIndex        =   30
               Top             =   2175
               Width           =   825
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Density:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   10
               Left            =   4530
               TabIndex        =   29
               Top             =   2670
               Width           =   720
            End
            Begin VB.Label lblL2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hf reference:"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   11
               Left            =   4530
               TabIndex        =   28
               Top             =   3150
               Width           =   1140
            End
         End
      End
      Begin VB.PictureBox PicHeader 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   30
         ScaleHeight     =   37
         ScaleMode       =   0  'User
         ScaleWidth      =   526.173
         TabIndex        =   23
         Top             =   30
         Width           =   8280
      End
   End
   Begin MMPlayerXProject.Button btnExit 
      Height          =   120
      Left            =   0
      TabIndex        =   96
      Top             =   0
      Width           =   225
      _extentx        =   397
      _extenty        =   212
      style           =   1
      buttoncolor     =   12632256
      mousepointer    =   99
   End
End
Attribute VB_Name = "frmDSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cWindows As New cWindowSkin
Dim icurrentFrame As Integer
Dim lbackcolor1 As Long
Dim lbackcolor2 As Long
Dim lpushedcolor As Long


Private Sub btnExit_Click()
    frmPopUp.mnuDSP_Click
End Sub

Private Sub chkDSP_Change(Index As Integer, Value As CheckBoxConstants)
    Dim i As Integer
    '=================================================================
    'Nota: Enable Many effects will cause performance
    '     CPU increases because they occur at the hardware level
    '     In particular, the equalizer (seldom)
    '=================================================================
    If Index = 7 Then frmMain.Button(16).Selected = chkDSP(Index).Value    ': Exit Sub
    FX_Disable   'Disable all FX
    'Enable FX that are checked
    For i = 0 To chkDSP.Count - 1
        If (chkDSP(i).Value = vbChecked) Then
            FX_Enable CLng(i)
            SetFX i
        End If

    Next i
End Sub

Private Sub Form_DblClick()
'Unload Me
End Sub

Private Sub Form_Load()
    loadSliders
    SetSliderValues
    LoadSkin
    loadConfiguration
    Const flag As Long = SWP_NOMOVE Or SWP_NOSIZE
    Call vkBtn_Click(icurrentFrame)
    'boolDspShow = True
    boolDspLoaded = True
    If OpcionesMusic.SiempreTop = True Then
        SetWindowPos frmDSP.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
    Else
        SetWindowPos frmDSP.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
    End If

End Sub
Public Sub LoadSkin()
    Dim ListScroll As New vkPrivateScroll
    Dim k
    On Error Resume Next
    Me.Height = Read_INI("FORM", "formheight", 6020, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    Me.Width = Read_INI("FORM", "formwidth", 8888, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")

    Set cWindows.FormularioPadre = Me
    cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
    cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\", True

    frameBack(0).Top = Read_INI("CONTAINER", "top", 12, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    frameBack(0).Left = Read_INI("CONTAINER", "left", 12, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    PicHeader.BackColor = frmPLST.picList.BackColor

    k = Read_Config_Button(btnExit, "Configuration", "exitButton", "0,0,10,10", tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    Set btnExit.PictureNormal = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\Exitnormal.bmp")
    Set btnExit.PictureOver = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\Exitover.bmp")
    Set btnExit.PictureDown = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\ExitDown.bmp")

    Set sldL2(0).PictureBack = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\HSliderBack.bmp")
    Set sldL2(0).PictureProgress = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\HSliderBack1.bmp")
    Set sldL2(0).bar = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\HbarNormal.bmp")
    Set sldL2(0).BarOver = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\HbarOver.bmp")
    Set sldL2(0).BarDown = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\HbarDown.bmp")

    For i = 1 To 11
        Set sldL2(i).PictureBack = sldL2(0).PictureBack
        Set sldL2(i).PictureProgress = sldL2(0).PictureProgress
        Set sldL2(i).bar = sldL2(0).bar
        Set sldL2(i).BarOver = sldL2(0).BarOver
        Set sldL2(i).BarDown = sldL2(0).BarDown
    Next

    For i = 0 To 6
        Set sldChorus(i).PictureBack = sldL2(0).PictureBack
        Set sldChorus(i).PictureProgress = sldL2(0).PictureProgress
        Set sldChorus(i).bar = sldL2(0).bar
        Set sldChorus(i).BarOver = sldL2(0).BarOver
        Set sldChorus(i).BarDown = sldL2(0).BarDown
    Next

    For i = 0 To 3
        Set sldWaves(i).PictureBack = sldL2(0).PictureBack
        Set sldWaves(i).PictureProgress = sldL2(0).PictureProgress
        Set sldWaves(i).bar = sldL2(0).bar
        Set sldWaves(i).BarOver = sldL2(0).BarOver
        Set sldWaves(i).BarDown = sldL2(0).BarDown
    Next

    For i = 0 To 5
        Set sldComp(i).PictureBack = sldL2(0).PictureBack
        Set sldComp(i).PictureProgress = sldL2(0).PictureProgress
        Set sldComp(i).bar = sldL2(0).bar
        Set sldComp(i).BarOver = sldL2(0).BarOver
        Set sldComp(i).BarDown = sldL2(0).BarDown
    Next

    For i = 0 To 4
        Set sldDis(i).PictureBack = sldL2(0).PictureBack
        Set sldDis(i).PictureProgress = sldL2(0).PictureProgress
        Set sldDis(i).bar = sldL2(0).bar
        Set sldDis(i).BarOver = sldL2(0).BarOver
        Set sldDis(i).BarDown = sldL2(0).BarDown
    Next

    For i = 0 To 4
        Set sldEcho(i).PictureBack = sldL2(0).PictureBack
        Set sldEcho(i).PictureProgress = sldL2(0).PictureProgress
        Set sldEcho(i).bar = sldL2(0).bar
        Set sldEcho(i).BarOver = sldL2(0).BarOver
        Set sldEcho(i).BarDown = sldL2(0).BarDown
    Next

    For i = 0 To 6
        Set sldFlan(i).PictureBack = sldL2(0).PictureBack
        Set sldFlan(i).PictureProgress = sldL2(0).PictureProgress
        Set sldFlan(i).bar = sldL2(0).bar
        Set sldFlan(i).BarOver = sldL2(0).BarOver
        Set sldFlan(i).BarDown = sldL2(0).BarDown
    Next

    For i = 0 To 1
        Set sldGarg(i).PictureBack = sldL2(0).PictureBack
        Set sldGarg(i).PictureProgress = sldL2(0).PictureProgress
        Set sldGarg(i).bar = sldL2(0).bar
        Set sldGarg(i).BarOver = sldL2(0).BarOver
        Set sldGarg(i).BarDown = sldL2(0).BarDown
    Next

    Set sldEQ(0).PictureBack = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\VSliderBack.bmp")
    Set sldEQ(0).PictureProgress = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\VSliderBack1.bmp")
    Set sldEQ(0).bar = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\VbarNormal.bmp")
    Set sldEQ(0).BarOver = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\VbarOver.bmp")
    Set sldEQ(0).BarDown = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\VbarDown.bmp")

    For i = 1 To 9
        Set sldEQ(i).PictureBack = sldEQ(0).PictureBack
        Set sldEQ(i).PictureProgress = sldEQ(0).PictureProgress
        Set sldEQ(i).bar = sldEQ(0).bar
        Set sldEQ(i).BarOver = sldEQ(0).BarOver
        Set sldEQ(i).BarDown = sldEQ(0).BarDown
    Next

    For i = 0 To 5
        frameBack(i).BackColor1 = cRead_INI("CONTAINER", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        frameBack(i).BackColor2 = cRead_INI("CONTAINER", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    Next

    For i = 0 To 3
        vkBtnEQ(i).UnRefreshControl = True
        vkBtnEQ(i).BackColor1 = cRead_INI("Button", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtnEQ(i).BackColor2 = cRead_INI("Button", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtnEQ(i).ForeColor = cRead_INI("BUTTON", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtnEQ(i).BorderColor = cRead_INI("BUTTON", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtnEQ(i).BackColorPushed1 = cRead_INI("Button", "BackColorPushed1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtnEQ(i).BackColorPushed2 = cRead_INI("Button", "BackColorPushed2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtnEQ(i).UnRefreshControl = False
        vkBtnEQ(i).Refresh
    Next

    lbackcolor1 = cRead_INI("Button", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    lbackcolor2 = cRead_INI("Button", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    lpushedcolor = cRead_INI("Button", "BackColorPushed2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    For i = 0 To 6
        vkBtn(i).UnRefreshControl = True
        vkBtn(i).BackColor1 = lbackcolor1
        vkBtn(i).BackColor2 = lbackcolor2
        vkBtn(i).ForeColor = cRead_INI("BUTTON", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtn(i).BorderColor = cRead_INI("BUTTON", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtn(i).BackColorPushed1 = cRead_INI("Button", "BackColorPushed1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtn(i).BackColorPushed2 = cRead_INI("Button", "BackColorPushed2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        vkBtn(i).UnRefreshControl = False
        vkBtn(i).Refresh
    Next

    For i = 0 To 4
        PicContainer(i).BackColor = cRead_INI("CONTAINER", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    Next

    txtEQPreset.BackColor = cRead_INI("TEXT", "backcolor", RGB(0, 60, 115), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    txtEQPreset.ForeColor = cRead_INI("TEXT", "forecolor", RGB(0, 60, 115), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")

    For i = 0 To 47
        lblL2(i).ForeColor = cRead_INI("LABEL", "forecolor", RGB(0, 226, 0), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    Next

    For i = 0 To 8
        chkDSP(i).ForeColor = cRead_INI("CHECKBOX", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
        chkDSP(i).BackColor = PicContainer(i).BackColor
    Next

    For i = 0 To 2
        vkCheckEqmode(i).ForeColor = cRead_INI("CHECKBOX", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    Next
    vklistEQ.UnRefreshControl = True
    ListScroll.BackColor = cRead_INI("LIST", "ScrollbackColor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    ListScroll.FrontColor = cRead_INI("LIST", "ScrollfrontColor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    ListScroll.BorderColor = cRead_INI("LIST", "ScrollborderColor", RGB(0, 0, 0), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    vklistEQ.ForeColor = cRead_INI("LIST", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    vklistEQ.BorderColor = cRead_INI("LIST", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    vklistEQ.BackColor = cRead_INI("LIST", "backcolor", RGB(70, 100, 50), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    vklistEQ.SelColor = cRead_INI("LIST", "selcolor", RGB(70, 100, 50), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\config.ini")
    vklistEQ.VScroll = ListScroll
    vklistEQ.UnRefreshControl = False
    vklistEQ.Refresh

    PrepareEQscale
    Call vkBtn_Click(0)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, 161, 2, 0
End Sub

Public Sub loadSliders()
    On Error Resume Next
    Dim i As Integer

    For i = 1 To 9
        Load sldEQ(i)
        sldEQ(i).Left = sldEQ(0).Left + 25 * i
        sldEQ(i).Visible = True
        lblEQ(0).Caption = "16kHz"
        PicContainer(0).CurrentX = sldEQ(i).Left + sldEQ(i).Width / 2 - lblEQ(0).Width / 2 + 1
        PicContainer(0).CurrentY = lblEQ(0).Top
        PicContainer(0).Print "8kHz"
    Next

    For i = 1 To 11
        Load sldL2(i)
        sldL2(i).Top = lblL2(i).Top + 3
        sldL2(i).Visible = True
        If i >= 6 Then sldL2(i).Left = lblL2(i).Left + 100
    Next

    For i = 1 To 6
        Load sldChorus(i)
        sldChorus(i).Top = lblL2(i + 12).Top + 3
        sldChorus(i).Visible = True
    Next

    For i = 1 To 3
        Load sldWaves(i)
        sldWaves(i).Top = lblL2(i + 19).Top + 3
        sldWaves(i).Visible = True
    Next

    For i = 1 To 5
        Load sldComp(i)
        sldComp(i).Top = lblL2(i + 23).Top + 3
        sldComp(i).Visible = True
    Next

    For i = 1 To 4
        Load sldDis(i)
        sldDis(i).Top = lblL2(i + 29).Top + 3
        sldDis(i).Visible = True
    Next

    For i = 1 To 4
        Load sldEcho(i)
        sldEcho(i).Top = lblL2(i + 34).Top + 3
        sldEcho(i).Visible = True
    Next

    For i = 1 To 6
        Load sldFlan(i)
        sldFlan(i).Top = lblL2(i + 39).Top + 3
        sldFlan(i).Visible = True
    Next

    For i = 1 To 1
        Load sldGarg(i)
        sldGarg(i).Top = lblL2(i + 46).Top + 3
        sldGarg(i).Visible = True
    Next

End Sub

Public Sub SetFX(intFX As Integer)
    Dim lngX As Long
    Select Case intFX
    Case 0
        FX_SetChorus sldChorus(0).Value, sldChorus(1).Value, sldChorus(2).Value, sldChorus(3).Value, sldChorus(4).Value, sldChorus(5).Value, sldChorus(6).Value
    Case 1
        FX_SetCompressor sldComp(0).Value, sldComp(1).Value, sldComp(2).Value, sldComp(3).Value, sldComp(4).Value, sldComp(5).Value
    Case 2
        FX_SetDistortion sldDis(0).Value, sldDis(1).Value, sldDis(2).Value, sldDis(3).Value, sldDis(4).Value
    Case 3
        FX_SetEcho sldEcho(0).Value, sldEcho(1).Value, sldEcho(2).Value, sldEcho(3).Value, sldEcho(4).Value
    Case 4
        FX_SetFlanger sldFlan(0).Value, sldFlan(1).Value, sldFlan(2).Value, sldFlan(3).Value, sldFlan(4).Value, sldFlan(5).Value, sldFlan(6).Value
    Case 5
        FX_SetGargle sldGarg(0).Value, sldGarg(1).Value
    Case 6
        FX_SetI3DL2Reverb sldL2(0).Value, sldL2(1).Value, sldL2(2).Value, sldL2(3).Value, sldL2(4).Value, sldL2(5).Value, sldL2(6).Value, sldL2(7).Value, sldL2(8).Value, sldL2(9).Value, sldL2(10).Value, sldL2(11).Value
    Case 7
        For i = 0 To 9 Step 1
            FX_SetEQ CLng(i), CLng(sldEQ(i).Value)
        Next
    Case 8
        FX_SetWavesReverb sldWaves(0).Value, sldWaves(1).Value, sldWaves(2).Value, sldWaves(3).Value
    End Select
End Sub
Sub Save_Equalizers()
    Dim Fnum As Integer, j As Integer
    Dim ArchivoINI As String
    Dim intClave As Integer

    'On Error GoTo BITCH
    ArchivoINI = tAppConfig.AppConfig & "Settings\Equalizer.eql"

    If Dir(ArchivoINI) <> "" Then    '// delete the file exists
        SetAttr ArchivoINI, vbNormal
        Kill ArchivoINI
    End If
    Fnum = FreeFile  '// random number to assign to the file
    Open ArchivoINI For Output As Fnum
    Print #Fnum, "+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+"
    Print #Fnum, "   EQUALIZER VALUES FOR MAHESHMP3 Player"
    Print #Fnum, "+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+"
    Print #Fnum, ""


    Dim i As Long
    Dim Arr As Variant
    Dim tIt As vkListItem
    ' since byref uses pointers we will hwve to get item as passing it instead of using functio item
    frmMain.ListEq.Clear
    For i = 1 To vklistEQ.ListCount
        Call vklistEQ.GetItem(tIt, i)
        Print #Fnum, "[equalizer_" & i - 1 & "]"
        Print #Fnum, "name=" & tIt.Text
        frmMain.ListEq.AddItem tIt.Text + "." + tIt.Key
        Arr = Split(tIt.Key, ",")
        For j = 0 To 9
            Print #Fnum, "eq" & j & "=" & str(Arr(j))
        Next j
    Next i
    Close Fnum
    Exit Sub
BITCH:
    MsgBox err.Description
End Sub
Private Sub sldeq_Chbange(Index As Integer, Value As Long)
    Dim k, i
    k = Abs(sldEQ(Index).Value - sldEQ(Index))
    For i = 1 To k
        If (Index + i) < 10 Then
            If (sldEQ(Index + i) + k - i) < 15 Then
                sldEQ(Index + i).Value = sldEQ(Index + i) + k - i
            Else
                sldEQ(Index + i).Value = 14
            End If
        End If
        If (Index - i) > 0 Then
            If (sldEQ(Index - i) + k - i) > 0 Then
                sldEQ(Index - i).Value = sldEQ(Index - i) + k - i
            Else
                sldEQ(Index - i).Value = 14
            End If
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveDspConfigurations
End Sub

Private Sub PicContainer_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> 0 Then Exit Sub
    Dim newval As Integer
    If Button = 1 Then
        activeSlider = -1
        Dim i
        For i = 0 To 9
            If sldEQ(i).Left < X And sldEQ(i).Left + sldEQ(i).Width > X Then
                activeSlider = i
                Exit For

            End If
        Next
        If activeSlider = -1 Then Exit Sub

        If Y <= sldEQ(0).Top And sldEQ(activeSlider).Value <> 10 Then sldEQ(activeSlider).Value = 10: Exit Sub
        sldEQ(activeSlider).Value = -10 + Min(20, ((sldEQ(0).Height + sldEQ(0).Top - Y) * 21#) / sldEQ(0).Height)
    End If
End Sub

Private Sub sldChorus_Change(Index As Integer, Value As Long)
    SetFX 0
End Sub

Private Sub sldComp_Change(Index As Integer, Value As Long)
    SetFX 1
End Sub

Private Sub sldDis_Change(Index As Integer, Value As Long)
    SetFX 2
End Sub

Private Sub sldEcho_Change(Index As Integer, Value As Long)
    SetFX 3
End Sub


Private Sub sldFlan_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    SetFX 4
End Sub

Private Sub sldGarg_Change(Index As Integer, Value As Long)
    SetFX 5
End Sub

Private Sub sldL2_Change(Index As Integer, Value As Long)
    SetFX 6
End Sub

Private Sub sldEQ_Change(Index As Integer, Value As Long)
    frmMain.Eq_SliderCtrl(Index).Value = Value
End Sub

Private Sub sldWaves_Change(Index As Integer, Value As Long)
    SetFX 8
End Sub

Public Sub vkBtn_Click(Index As Integer)

    If Index = icurrentFrame Then Exit Sub
    If Index = 5 Then DSPReset: Exit Sub
    If Index = 6 Then DSPClear: Exit Sub
    frameBack(Index + 1).ZOrder vbBringToFront
    icurrentFrame = Index

    For i = 0 To vkBtn.UBound
        vkBtn(i).BackColor1 = lbackcolor1
        vkBtn(i).BackColor2 = lbackcolor2
    Next
    vkBtn(Index).BackColor1 = lpushedcolor
    vkBtn(Index).BackColor2 = lpushedcolor
    DrawLine (vkBtn(Index).Left / Screen.TwipsPerPixelX) + 1, 0, (vkBtn(Index).Left + vkBtn(Index).Width) / Screen.TwipsPerPixelX - 1, 0, frameBack(Index + 1).hDC, lpushedcolor
End Sub

Private Sub vkBtnEQ_Click(Index As Integer)
    Dim i As Long
    Select Case Index

    Case 0:

    Case 1:
        Dim T, s As String
        Dim j As Integer
        T = ""
        If txtEQPreset.Text <> "" Then
            s = txtEQPreset.Text
        Else
            s = "User Defined"
        End If

        For i = 0 To 9
            T = T + str(sldEQ(i).Value) + ","
        Next i
        vklistEQ.AddItem s, , T
        vklistEQ.Refresh
        Save_Equalizers
    Case 2:
        i = 1
        While (vklistEQ.CheckCount > 0)
            If vklistEQ.Item(i).Checked = True Then
                vklistEQ.RemoveItem (i)
            Else
                i = i + 1
            End If
        Wend
        'vklistEQ.UnCheckAll
        vklistEQ.Refresh
        Save_Equalizers
    Case 3:
        vklistEQ.Clear
        vklistEQ.UnRefreshControl = True
        vklistEQ.AddItem "Rock", , "6,7,5,3,2,1,-1,0,3,5"
        vklistEQ.AddItem "Pop", , "2,7,5,6,4,1,-1,-1,0,1"
        vklistEQ.AddItem "Jazz", , "6,7,5,3,2,1,-1,0,3,5"
        vklistEQ.AddItem "Classical", , "0,0,0,0,0,0,-2,-2,-2,3"
        vklistEQ.AddItem "Vocal", , "6,7,5,3,2,1,-1,0,3,5,8"
        vklistEQ.AddItem "Full Bass", , "6,7,5,3,2,1,-1,0,3,5,8"
        vklistEQ.AddItem "Full Treble", , "6,7,5,3,2,1,-1,0,3,5,8"
        vklistEQ.AddItem "Headphones", , "4,5,3,1,0,-1,-2,-1,5,8"
        vklistEQ.AddItem "Live", , "-2,-1,0,1,2,1,2,2,1,1"
        vklistEQ.AddItem "Party", , "3,3,0,0,0,0,0,0,3,4"
        vklistEQ.AddItem "Soft", , "3,4,2,1,0,-1,-2,0,2,6"
        vklistEQ.AddItem "Reggae", , "0,0,-1,-2,0,2,3,0,0,0"
        vklistEQ.AddItem "Large Hall", , "8,7,5,5,2,1,-2,-1,0,1"
        vklistEQ.UnRefreshControl = False
        vklistEQ.Refresh
        Save_Equalizers
    End Select

End Sub

Private Sub DSPClear()
    For i = 0 To chkDSP.Count - 1
        chkDSP(i).Value = vbUnchecked
    Next i
End Sub

Private Sub DSPReset()
    Select Case icurrentFrame

    Case 1    '// ID3L2 Rev
        sldL2(0).Value = -1000: sldL2(1).Value = 0
        sldL2(2).Value = 0: sldL2(3).Value = 2
        sldL2(4).Value = 1: sldL2(5).Value = -2602
        sldL2(6).Value = 1: sldL2(7).Value = 200
        sldL2(8).Value = 0: sldL2(9).Value = 100
        sldL2(10).Value = 100: sldL2(11).Value = 5000

    Case 2    '// chorus + wave rev
        sldChorus(0).Value = 50: sldChorus(1).Value = 25
        sldChorus(2).Value = 0: sldChorus(3).Value = 0
        sldChorus(4).Value = 1: sldChorus(5).Value = 0
        sldChorus(0).Value = 0
        sldWaves(0).Value = 0: sldWaves(1).Value = 0
        sldWaves(2).Value = 1000: sldWaves(3).Value = 0

    Case 3    '// compressor+ distortion
        sldComp(0).Value = 0: sldComp(1).Value = 0
        sldComp(2).Value = 50: sldComp(3).Value = -10
        sldComp(4).Value = 10: sldComp(5).Value = 0
        sldDis(0).Value = 0: sldDis(1).Value = 50
        sldDis(2).Value = 4000: sldDis(3).Value = 4000
        sldDis(4).Value = 4000

    Case 4    '// echo+flanger+gargle
        sldEcho(0).Value = 95: sldEcho(1).Value = 0
        sldEcho(2).Value = 333: sldEcho(3).Value = 555
        sldEcho(4).Value = 0
        sldFlan(0).Value = 50: sldFlan(1).Value = 25
        sldFlan(2).Value = 0: sldFlan(3).Value = 0
        sldFlan(4).Value = 1: sldFlan(5).Value = 0
        sldFlan(6).Value = 0
        sldGarg(0).Value = 500: sldGarg(1).Value = 0

    End Select

End Sub





Private Sub vkCheckEqmode_Click(Index As Integer)
    Select Case Index
    Case 0:
        bEQFreeMove = Not bEQFreeMove
        frmMain.Button(17).Selected = bEQFreeMove
    End Select
End Sub

Private Sub vklistEQ_ItemClick(Item As vkListItem)
    If Item Is Nothing Then Exit Sub    ' error comes when clicked in clear region so 'nothing' item is returned
    Dim i As Integer
    Dim Arr As Variant
    'If item = Nothing Then Exit Sub
    txtEQPreset.Text = "PRESET: " + Item.Text
    Arr = Split(Item.Key, ",")
    For i = 0 To 9
        sldEQ(i).Value = CInt(Arr(i))
    Next

End Sub

Public Sub SetSliderValues()
'I3dl2 reverb settings
'room
    sldL2(0).Min = -10000
    sldL2(0).Max = 0
    'room hf
    sldL2(1).Min = -10000
    sldL2(1).Max = 0
    'roll off  factor
    sldL2(2).Min = 0
    sldL2(2).Max = 10
    'decay time
    sldL2(3).Min = 0
    sldL2(3).Max = 20
    'decay hf ratio
    sldL2(4).Min = 0
    sldL2(4).Max = 2
    'reflection
    sldL2(5).Min = -10000
    sldL2(5).Max = 10000
    'reflection delay
    sldL2(6).Min = 0
    sldL2(6).Max = 1
    'reverb
    sldL2(7).Min = -100002
    sldL2(7).Max = 2000
    'reverb delay
    sldL2(8).Min = 0
    sldL2(8).Max = 1
    'diffusion
    sldL2(9).Min = 0
    sldL2(9).Max = 100
    'density
    sldL2(10).Min = 0
    sldL2(10).Max = 100
    'Hf reference
    sldL2(11).Min = 20
    sldL2(11).Max = 20000

    'Chorus settings
    'wet dry mix
    sldChorus(0).Min = 0
    sldChorus(0).Max = 100
    'depth
    sldChorus(1).Min = 0
    sldChorus(1).Max = 100
    'feedback
    sldChorus(2).Min = -99
    sldChorus(2).Max = 99
    'frequency
    sldChorus(3).Min = 0
    sldChorus(3).Max = 10
    'waveform
    sldChorus(4).Min = 0
    sldChorus(4).Max = 1
    'delay
    sldChorus(5).Min = 5
    sldChorus(5).Max = 195
    'phase
    sldChorus(6).Min = 0
    sldChorus(6).Max = 4

    'Wave reverb settings
    'in gain
    sldWaves(0).Min = -96
    sldWaves(0).Max = 0
    'reverb mix
    sldWaves(1).Min = -96
    sldWaves(1).Max = 0
    'reverb time
    sldWaves(2).Min = 0
    sldWaves(2).Max = 3000
    'high frequency ratio
    sldWaves(3).Min = 0
    sldWaves(3).Max = 1

    'compression settings
    'gain
    sldComp(0).Min = -60
    sldComp(0).Max = 60
    'attack
    sldComp(1).Min = -60
    sldComp(1).Max = 60
    'release
    sldComp(2).Min = 50
    sldComp(2).Max = 3000
    'threshold
    sldComp(3).Min = -60
    sldComp(3).Max = 0
    'ratio
    sldComp(4).Min = 1
    sldComp(4).Max = 100
    'pre delay
    sldComp(5).Min = 0
    sldComp(5).Max = 4

    'Distortion settings
    'gain
    sldDis(0).Min = -60
    sldDis(0).Max = 0
    'edge
    sldDis(1).Min = 0
    sldDis(1).Max = 100
    'eq center freq
    sldDis(2).Min = 100
    sldDis(2).Max = 8000
    'geq bandwidth
    sldDis(3).Min = 100
    sldDis(3).Max = 8000
    'low pass filter
    sldDis(4).Min = 100
    sldDis(4).Max = 8000


    'echo settings
    'wet dry mix
    sldEcho(0).Min = 0
    sldEcho(0).Max = 100
    'feedback
    sldEcho(1).Min = 0
    sldEcho(1).Max = 100
    'left delay
    sldEcho(2).Min = 1
    sldEcho(2).Max = 2000
    'right delay
    sldEcho(3).Min = 1
    sldEcho(3).Max = 2000
    'pan delay
    sldEcho(4).Min = 0
    sldEcho(4).Max = 1

    'Flanger settings
    'wet dry mix
    sldFlan(0).Min = 0
    sldFlan(0).Max = 100
    'depth
    sldFlan(1).Min = 0
    sldFlan(1).Max = 100
    'feedback
    sldFlan(2).Min = -99
    sldFlan(2).Max = 99
    'frequency
    sldFlan(3).Min = 0
    sldFlan(3).Max = 10
    'waveform
    sldFlan(4).Min = 0
    sldFlan(4).Max = 1
    'delay
    sldFlan(5).Min = 0
    sldFlan(5).Max = 4
    'phase
    sldFlan(6).Min = 0
    sldFlan(6).Max = 4


    'Gargle settings
    'Hz
    sldGarg(0).Min = 1
    sldGarg(0).Max = 1000
    'waveshape
    sldGarg(1).Min = 0
    sldGarg(1).Max = 1

End Sub

Public Sub loadConfiguration()
    Dim strrs
    On Error Resume Next
    With frmDSP


        DoEvents
        strRes = Read_INI("Equalizer", "Preset", -1, , True)
        If strRes >= 0 Or strRes <= .vklistEQ.ListCount - 1 Then .vklistEQ.ListIndex = CInt(strRes)

        '===============================================================================
        ' SOUND EFFECTS
        strRes = Read_INI("Sound_Effect", "Chorus", 0, , True)
        If CBool(strRes) = True Then .chkDSP(0).Value = 1

        For i = 0 To .sldChorus.Count - 1
            strRes = Read_INI("Sound_Effect", "Chorus_" & i, 0, , True)
            .sldChorus(i).Value = CInt(strRes)
        Next i

        strRes = Read_INI("Sound_Effect", "Compressor", 0, , True)
        If CBool(strRes) = True Then .chkDSP(1).Value = 1

        For i = 0 To .sldComp.Count - 1
            strRes = Read_INI("Sound_Effect", "Compressor_" & i, 0, , True)
            .sldComp(i).Value = CInt(strRes)
        Next i

        strRes = Read_INI("Sound_Effect", "Distortion", 0, , True)
        If CBool(strRes) = True Then .chkDSP(2).Value = 1

        For i = 0 To .sldDis.Count - 1
            strRes = Read_INI("Sound_Effect", "Distortion_" & i, 0, , True)
            .sldDis(i).Value = CInt(strRes)
        Next i

        strRes = Read_INI("Sound_Effect", "Echo", 0, , True)
        If CBool(strRes) = True Then .chkDSP(3).Value = 1

        For i = 0 To .sldEcho.Count - 1
            strRes = Read_INI("Sound_Effect", "Echo_" & i, 0, , True)
            .sldEcho(i).Value = CInt(strRes)
        Next i

        strRes = Read_INI("Sound_Effect", "Flanger", 0, , True)
        If CBool(strRes) = True Then .chkDSP(4).Value = 1

        For i = 0 To .sldFlan.Count - 1
            strRes = Read_INI("Sound_Effect", "Flanger_" & i, 0, , True)
            .sldFlan(i).Value = CInt(strRes)
        Next i

        strRes = Read_INI("Sound_Effect", "Gargle", 0, , True)
        If CBool(strRes) = True Then .chkDSP(5).Value = 1

        For i = 0 To .sldGarg.Count - 1
            strRes = Read_INI("Sound_Effect", "Gargle_" & i, 0, , True)
            .sldGarg(i).Value = CInt(strRes)
        Next i

        strRes = Read_INI("Sound_Effect", "L2Reverb", 0, , True)
        If CBool(strRes) = True Then .chkDSP(6).Value = 1

        For i = 0 To .sldL2.Count - 1
            strRes = Read_INI("Sound_Effect", "L2Reverb_" & i, 0, , True)
            .sldL2(i).Value = CInt(strRes)
        Next i

        strRes = Read_INI("Sound_Effect", "WReverb", 0, , True)
        If CBool(strRes) = True Then .chkDSP(8).Value = 1

        For i = 0 To .sldWaves.Count - 1
            strRes = Read_INI("Sound_Effect", "WReverb_" & i, 0, , True)
            .sldWaves(i).Value = CInt(strRes)
        Next i
        For i = 0 To 9
            sldEQ(i).Value = frmMain.Eq_SliderCtrl(i).Value
        Next
        If frmMain.Button(16).Selected Then chkDSP(7).Value = vbChecked

    End With

    Dim sValue
    Dim s, T As String
    i = 0
    'DoEvents
    DoEvents
    Do
        T = ""
        s = ""
        s = Read_INI("equalizer_" & i, "name", "", , , tAppConfig.AppConfig & "Settings\Equalizer.eql")
        For j = 0 To 9
            sValue = Read_INI("equalizer_" & i, "eq" & j, 0, , , tAppConfig.AppConfig & "Settings\Equalizer.eql")
            If IsNumeric(sValue) Then T = T + str(sValue) + ","
        Next j
        If s <> "" And T <> "" Then vklistEQ.AddItem s, , T
        i = i + 1
    Loop While s <> ""

    If bEQFreeMove Then vkCheckEqmode(0).Value = vbChecked
    strRes = Read_INI("Configuration", "PlayDFXStarting", 1, , True)
    If CBool(strRes) = True Then bPlayDFXStarting = True
    If bPlayDFXStarting = True Then Call DSPClear    'turn off dsp
    icurrentFrame = 1

End Sub

Public Sub SaveDspConfigurations()
    Dim ArchivoINI As String
    Dim INIcheck As String
    Dim Fnum As Integer
    ArchivoINI = tAppConfig.AppPath & App.EXEName & ".ini"

    '// Chekar los atributos
    INIcheck = Dir(ArchivoINI, vbNormal + vbSystem + vbHidden + vbReadOnly + vbArchive)

    '// Si no se encuentra hacerlo...
    If INIcheck = "" Then
        Fnum = FreeFile  '// numeroaleatorio para asignar al archivo
        Open ArchivoINI For Output As Fnum
        Close
        'SetAttr ArchivoINI, vbHidden + vbSystem
    End If
    With frmDSP
        Write_INI "Equalizer", "Preset", .vklistEQ.ListIndex, ArchivoINI

        Write_INI "Sound_Effect", "Chorus", CBool(.chkDSP(0).Value), ArchivoINI
        For i = 0 To .sldChorus.Count - 1
            Write_INI "Sound_Effect", "Chorus_" & i, .sldChorus(i).Value, ArchivoINI
        Next i

        Write_INI "Sound_Effect", "Compressor", CBool(.chkDSP(1).Value), ArchivoINI
        For i = 0 To .sldComp.Count - 1
            Write_INI "Sound_Effect", "Compressor_" & i, .sldComp(i).Value, ArchivoINI
        Next i

        Write_INI "Sound_Effect", "Distortion", CBool(.chkDSP(2).Value), ArchivoINI
        For i = 0 To .sldDis.Count - 1
            Write_INI "Sound_Effect", "Distortion_" & i, .sldDis(i).Value, ArchivoINI
        Next i

        Write_INI "Sound_Effect", "Echo", CBool(.chkDSP(3).Value), ArchivoINI
        For i = 0 To .sldEcho.Count - 1
            Write_INI "Sound_Effect", "Echo_" & i, .sldEcho(i).Value, ArchivoINI
        Next i

        Write_INI "Sound_Effect", "Flanger", CBool(.chkDSP(4).Value), ArchivoINI
        For i = 0 To .sldFlan.Count - 1
            Write_INI "Sound_Effect", "Flanger_" & i, .sldFlan(i).Value, ArchivoINI
        Next i

        Write_INI "Sound_Effect", "Gargle", CBool(.chkDSP(5).Value), ArchivoINI
        For i = 0 To .sldGarg.Count - 1
            Write_INI "Sound_Effect", "Gargle_" & i, .sldGarg(i).Value, ArchivoINI
        Next i

        Write_INI "Sound_Effect", "L2Reverb", CBool(.chkDSP(6).Value), ArchivoINI
        For i = 0 To .sldL2.Count - 1
            Write_INI "Sound_Effect", "L2Reverb_" & i, .sldL2(i).Value, ArchivoINI
        Next i

        Write_INI "Sound_Effect", "WReverb", CBool(.chkDSP(8).Value), ArchivoINI
        For i = 0 To .sldWaves.Count - 1
            Write_INI "Sound_Effect", "WReverb_" & i, .sldWaves(i).Value, ArchivoINI
        Next i

        Write_INI "Sound_Effect", "PlayDFXStarting", bPlayDFXStarting, ArchivoINI

    End With

End Sub

Public Sub PrepareEQscale()
    For i = 0 To 20
        PicContainer(0).Line (sldEQ(0).Left - 9, sldEQ(0).Top + i * sldEQ(0).Height / 20#)-(sldEQ(0).Left - 6, sldEQ(0).Top + i * sldEQ(0).Height / 20#)
        PicContainer(0).Line (sldEQ(9).Left + sldEQ(9).Width + 6, sldEQ(0).Top + i * sldEQ(0).Height / 20#)-(sldEQ(9).Left + sldEQ(9).Width + 9, sldEQ(0).Top + i * sldEQ(0).Height / 20#)
    Next
    PicContainer(0).Line (sldEQ(0).Left - 13, sldEQ(0).Top)-(sldEQ(0).Left - 6, sldEQ(0).Top)
    PicContainer(0).Line (sldEQ(0).Left - 11, sldEQ(0).Top + 10 * sldEQ(0).Height / 20#)-(sldEQ(0).Left - 6, sldEQ(0).Top + 10 * sldEQ(0).Height / 20#)
    PicContainer(0).Line (sldEQ(0).Left - 13, sldEQ(0).Top + sldEQ(0).Height)-(sldEQ(0).Left - 6, sldEQ(0).Top + sldEQ(0).Height)
    PicContainer(0).Line (sldEQ(9).Left + sldEQ(9).Width + 6, sldEQ(0).Top)-(sldEQ(9).Left + sldEQ(9).Width + 13, sldEQ(0).Top)
    PicContainer(0).Line (sldEQ(9).Left + sldEQ(9).Width + 6, sldEQ(0).Top + 10 * sldEQ(0).Height / 20#)-(sldEQ(9).Left + sldEQ(9).Width + 11, sldEQ(0).Top + 10 * sldEQ(0).Height / 20#)
    PicContainer(0).Line (sldEQ(9).Left + sldEQ(9).Width + 6, sldEQ(0).Top + sldEQ(0).Height)-(sldEQ(9).Left + sldEQ(9).Width + 13, sldEQ(0).Top + sldEQ(0).Height)
    PicContainer(0).CurrentY = sldEQ(0).Top - 4
    PicContainer(0).CurrentX = sldEQ(0).Left - PicContainer(0).TextWidth("+10dB") - 18
    PicContainer(0).Print "+10dB"
    PicContainer(0).CurrentY = sldEQ(0).Top - 4
    PicContainer(0).CurrentX = sldEQ(9).Left + sldEQ(0).Width + 17
    PicContainer(0).Print "+10dB"

    PicContainer(0).CurrentY = sldEQ(0).Top + sldEQ(0).Height - 4
    PicContainer(0).CurrentX = sldEQ(0).Left - PicContainer(0).TextWidth("+10dB") - 18
    PicContainer(0).Print "-10dB"
    PicContainer(0).CurrentY = sldEQ(0).Top + sldEQ(0).Height - 5
    PicContainer(0).CurrentX = sldEQ(9).Left + sldEQ(0).Width + 17
    PicContainer(0).Print "-10dB"
End Sub

