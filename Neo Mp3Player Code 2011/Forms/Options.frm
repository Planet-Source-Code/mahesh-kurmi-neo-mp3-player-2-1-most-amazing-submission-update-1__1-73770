VERSION 5.00
Begin VB.Form frmOptions 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   " Options"
   ClientHeight    =   10680
   ClientLeft      =   2505
   ClientTop       =   3885
   ClientWidth     =   16140
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   712
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1076
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picscreenshot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   -1800
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   205
      TabIndex        =   29
      Top             =   9180
      Visible         =   0   'False
      Width           =   3075
   End
   Begin MMPlayerXProject.vkFrame FrameContainer 
      Height          =   9750
      Left            =   750
      TabIndex        =   1
      Top             =   60
      Width           =   13815
      _extentx        =   24368
      _extenty        =   17198
      font            =   "Options.frx":000C
      showtitle       =   0   'False
      bordercolor     =   0
      roundangle      =   0
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3915
         Index           =   4
         Left            =   6810
         TabIndex        =   20
         Top             =   7920
         Width           =   6480
         _extentx        =   11430
         _extenty        =   6906
         font            =   "Options.frx":0034
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin MMPlayerXProject.vkCheck vkCheckIcon 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   0
            Left            =   2610
            TabIndex        =   46
            Top             =   2575
            Width           =   1965
            _extentx        =   3466
            _extenty        =   397
            backstyle       =   0
            caption         =   "Show as Tray Icon"
            font            =   "Options.frx":005C
         End
         Begin MMPlayerXProject.vkCheck vkCheckIcon 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   6
            Left            =   2610
            TabIndex        =   45
            Top             =   2910
            Width           =   1845
            _extentx        =   3254
            _extenty        =   397
            backstyle       =   0
            caption         =   "Show in Taskbar"
            font            =   "Options.frx":0084
         End
         Begin MMPlayerXProject.vkListBox vkListType 
            Height          =   3195
            Left            =   150
            TabIndex        =   36
            Top             =   460
            Width           =   2325
            _extentx        =   4101
            _extenty        =   5636
            backcolor       =   15722980
            bordercolor     =   0
            sorted          =   0
            stylecheckbox   =   -1  'True
            font            =   "Options.frx":00AC
            selcolor        =   4210752
            borderselcolor  =   8421504
            usedefautitemsettings=   -1  'True
         End
         Begin MMPlayerXProject.vkCheck vkCheckIcon 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   5
            Left            =   2610
            TabIndex        =   35
            Top             =   1800
            Width           =   1305
            _extentx        =   2302
            _extenty        =   397
            backstyle       =   0
            caption         =   "Next Icon"
            font            =   "Options.frx":00D4
         End
         Begin MMPlayerXProject.vkCheck vkCheckIcon 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   4
            Left            =   2610
            TabIndex        =   34
            Top             =   1470
            Width           =   1305
            _extentx        =   2302
            _extenty        =   397
            backstyle       =   0
            caption         =   "Stop Icon"
            font            =   "Options.frx":00FC
         End
         Begin MMPlayerXProject.vkCheck vkCheckIcon 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   3
            Left            =   2610
            TabIndex        =   33
            Top             =   1140
            Width           =   1335
            _extentx        =   2355
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Pause Icon"
            font            =   "Options.frx":0124
         End
         Begin MMPlayerXProject.vkCheck vkCheckIcon 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   1
            Left            =   2610
            TabIndex        =   32
            Top             =   450
            Width           =   1635
            _extentx        =   2884
            _extenty        =   397
            backstyle       =   0
            caption         =   "Previous Icon"
            font            =   "Options.frx":014C
         End
         Begin MMPlayerXProject.vkCheck vkCheckIcon 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   2
            Left            =   2610
            TabIndex        =   31
            Top             =   810
            Width           =   1215
            _extentx        =   2143
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Play Icon"
            font            =   "Options.frx":0174
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show system tray icon:"
            Height          =   195
            Index           =   14
            Left            =   2580
            TabIndex        =   47
            Top             =   2275
            Width           =   2025
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Register File types:"
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   37
            Top             =   180
            Width           =   1650
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show system tray icon:"
            Height          =   195
            Index           =   12
            Left            =   2610
            TabIndex        =   30
            Top             =   180
            Width           =   2025
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3915
         Index           =   3
         Left            =   7800
         TabIndex        =   21
         Top             =   3960
         Width           =   6300
         _extentx        =   11113
         _extenty        =   6906
         font            =   "Options.frx":019C
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin MMPlayerXProject.Eq_SliderCtrl sldSettings 
            Height          =   105
            Index           =   6
            Left            =   4560
            TabIndex        =   49
            Top             =   3300
            Width           =   1455
            _extentx        =   2566
            _extenty        =   185
            pictureback     =   "Options.frx":01C4
            pictureprogress =   "Options.frx":0A12
            bardown         =   "Options.frx":1260
            barover         =   "Options.frx":13CA
            bar             =   "Options.frx":1534
            backcolor       =   -2147483643
            min             =   2
            max             =   15
            value           =   5
            position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl sldSettings 
            Height          =   105
            Index           =   5
            Left            =   4560
            TabIndex        =   44
            Top             =   2850
            Width           =   1455
            _extentx        =   185
            _extenty        =   2566
            pictureback     =   "Options.frx":169E
            pictureprogress =   "Options.frx":1EEC
            bardown         =   "Options.frx":273A
            barover         =   "Options.frx":28A4
            bar             =   "Options.frx":2A0E
            backcolor       =   -2147483643
            min             =   3
            max             =   26
            value           =   13
            position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl sldSettings 
            Height          =   105
            Index           =   4
            Left            =   4560
            TabIndex        =   43
            Top             =   2400
            Width           =   1455
            _extentx        =   2566
            _extenty        =   185
            pictureback     =   "Options.frx":2B78
            pictureprogress =   "Options.frx":33C6
            bardown         =   "Options.frx":3C14
            barover         =   "Options.frx":3D7E
            bar             =   "Options.frx":3EE8
            backcolor       =   -2147483643
            min             =   20
            value           =   50
            position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl sldSettings 
            Height          =   105
            Index           =   3
            Left            =   4560
            TabIndex        =   42
            Top             =   1950
            Width           =   1455
            _extentx        =   2566
            _extenty        =   185
            pictureback     =   "Options.frx":4052
            pictureprogress =   "Options.frx":48A0
            bardown         =   "Options.frx":50EE
            barover         =   "Options.frx":5258
            bar             =   "Options.frx":53C2
            backcolor       =   -2147483643
            min             =   40
            max             =   250
            value           =   200
            position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl sldSettings 
            Height          =   105
            Index           =   2
            Left            =   4560
            TabIndex        =   41
            Top             =   1530
            Width           =   1455
            _extentx        =   2566
            _extenty        =   185
            pictureback     =   "Options.frx":552C
            pictureprogress =   "Options.frx":5D7A
            bardown         =   "Options.frx":65C8
            barover         =   "Options.frx":6732
            bar             =   "Options.frx":689C
            backcolor       =   -2147483643
            min             =   10
            value           =   10
            position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl sldSettings 
            Height          =   105
            Index           =   0
            Left            =   4560
            TabIndex        =   39
            Top             =   630
            Width           =   1455
            _extentx        =   2566
            _extenty        =   185
            pictureback     =   "Options.frx":6A06
            pictureprogress =   "Options.frx":7254
            bardown         =   "Options.frx":7AA2
            barover         =   "Options.frx":7C0C
            bar             =   "Options.frx":7D76
            backcolor       =   14737632
            max             =   400
            value           =   20
            position        =   1
         End
         Begin MMPlayerXProject.Eq_SliderCtrl sldSettings 
            Height          =   105
            Index           =   1
            Left            =   4560
            TabIndex        =   38
            Top             =   1080
            Width           =   1455
            _extentx        =   2566
            _extenty        =   185
            pictureback     =   "Options.frx":7EE0
            pictureprogress =   "Options.frx":872E
            bardown         =   "Options.frx":8F7C
            barover         =   "Options.frx":90E6
            bar             =   "Options.frx":9250
            backcolor       =   14737632
            max             =   400
            value           =   100
            position        =   1
         End
         Begin VB.Label lblsettings 
            BackStyle       =   0  'Transparent
            Caption         =   "Seek track in 1 mouse scroll (in s):"
            Height          =   225
            Index           =   6
            Left            =   300
            TabIndex        =   63
            Top             =   3240
            Width           =   3495
         End
         Begin VB.Label lblsettings 
            BackStyle       =   0  'Transparent
            Caption         =   "Seek Volume in 1 mouse scroll (in %):"
            Height          =   225
            Index           =   5
            Left            =   300
            TabIndex        =   62
            Top             =   2790
            Width           =   3705
         End
         Begin VB.Label lblsettings 
            BackStyle       =   0  'Transparent
            Caption         =   "Spectrum refresh rate (in frames/s):"
            Height          =   225
            Index           =   4
            Left            =   300
            TabIndex        =   61
            Top             =   2340
            Width           =   3495
         End
         Begin VB.Label lblsettings 
            BackStyle       =   0  'Transparent
            Caption         =   "Scrollling text Interval(in ms):"
            Height          =   225
            Index           =   3
            Left            =   300
            TabIndex        =   60
            Top             =   1860
            Width           =   3495
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Opacity(in %)"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   300
            TabIndex        =   59
            Top             =   1440
            Width           =   1185
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Crossfade in Stop/Pause (in ms):"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   58
            Top             =   990
            Width           =   2865
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Crossfade between tracks (in ms):"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   57
            Top             =   570
            Width           =   2985
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PLAYER SETTINGS:"
            Height          =   195
            Index           =   7
            Left            =   330
            TabIndex        =   40
            Top             =   210
            Width           =   1665
         End
      End
      Begin MMPlayerXProject.vkCommand CMD 
         Height          =   405
         Index           =   2
         Left            =   3540
         TabIndex        =   12
         Top             =   4080
         Width           =   1365
         _extentx        =   2408
         _extenty        =   714
         caption         =   "OK"
         font            =   "Options.frx":93BA
      End
      Begin MMPlayerXProject.vkCommand CMD 
         CausesValidation=   0   'False
         Height          =   405
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   4080
         Width           =   1845
         _extentx        =   3254
         _extenty        =   714
         caption         =   "Load Defaults"
         font            =   "Options.frx":93E2
      End
      Begin MMPlayerXProject.vkCommand CMD 
         Height          =   405
         Index           =   3
         Left            =   5010
         TabIndex        =   10
         Top             =   4080
         Width           =   1305
         _extentx        =   2302
         _extenty        =   714
         caption         =   "Apply"
         font            =   "Options.frx":940A
      End
      Begin MMPlayerXProject.vkCommand CMD 
         Height          =   405
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   4080
         Width           =   1395
         _extentx        =   2461
         _extenty        =   714
         caption         =   "Cancel"
         font            =   "Options.frx":9432
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3915
         Index           =   0
         Left            =   1080
         TabIndex        =   2
         Top             =   4560
         Width           =   6300
         _extentx        =   11113
         _extenty        =   6906
         font            =   "Options.frx":945A
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin MMPlayerXProject.vkOptionButton optPlstCase 
            Height          =   285
            Index           =   2
            Left            =   4650
            TabIndex        =   70
            Top             =   3540
            Width           =   1305
            _extentx        =   2302
            _extenty        =   503
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "UpperCase"
            font            =   "Options.frx":9482
            group           =   2
         End
         Begin MMPlayerXProject.vkOptionButton optPlstCase 
            Height          =   285
            Index           =   1
            Left            =   4650
            TabIndex        =   69
            Top             =   3270
            Width           =   1305
            _extentx        =   2302
            _extenty        =   503
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "LowerCase"
            font            =   "Options.frx":94AA
            value           =   1
            group           =   2
         End
         Begin MMPlayerXProject.vkOptionButton optPlstCase 
            Height          =   285
            Index           =   0
            Left            =   4650
            TabIndex        =   68
            Top             =   3000
            Width           =   1215
            _extentx        =   556
            _extenty        =   503
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Normal"
            font            =   "Options.frx":94D2
            group           =   2
         End
         Begin MMPlayerXProject.vkOptionButton optScrollType 
            Height          =   285
            Index           =   1
            Left            =   150
            TabIndex        =   67
            Top             =   3360
            Width           =   1065
            _extentx        =   1879
            _extenty        =   503
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "ZigZag"
            font            =   "Options.frx":94FA
            group           =   1
         End
         Begin MMPlayerXProject.vkOptionButton optScrollType 
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   66
            Top             =   3060
            Width           =   1215
            _extentx        =   556
            _extenty        =   503
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Rolling"
            font            =   "Options.frx":9522
            value           =   1
            group           =   1
         End
         Begin VB.TextBox txtDisplay 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   90
            TabIndex        =   5
            Text            =   "%S - %A (%T)"
            Top             =   930
            Width           =   5985
         End
         Begin VB.TextBox txtDisplay 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   4
            Text            =   "%A - %S"
            Top             =   300
            Width           =   5985
         End
         Begin VB.TextBox txtFormat 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1425
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   3
            Text            =   "Options.frx":954A
            Top             =   1320
            Width           =   5985
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Display tracks in"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   4650
            TabIndex        =   65
            Top             =   2820
            Width           =   1425
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "scroll text Format:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   90
            TabIndex        =   8
            Top             =   690
            Width           =   1575
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Playlist Format for New Entry"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   7
            Top             =   90
            Width           =   2490
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Scroll Type:"
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   11
            Left            =   150
            TabIndex        =   6
            Top             =   2820
            Width           =   1035
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3915
         Index           =   1
         Left            =   1050
         TabIndex        =   13
         Top             =   510
         Width           =   6300
         _extentx        =   11113
         _extenty        =   6906
         font            =   "Options.frx":96AB
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin VB.TextBox txtAppConfig 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   120
            TabIndex        =   17
            Top             =   300
            Width           =   5535
         End
         Begin VB.TextBox txtSkininfo 
            Appearance      =   0  'Flat
            Height          =   1005
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   2790
            Width           =   6045
         End
         Begin MMPlayerXProject.vkCheck chkUsefile 
            CausesValidation=   0   'False
            Height          =   225
            Left            =   120
            TabIndex        =   15
            Top             =   2520
            Width           =   2625
            _extentx        =   4630
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Load Skins from region file"
            font            =   "Options.frx":96D3
         End
         Begin MMPlayerXProject.vkCommand CMD 
            Height          =   315
            Index           =   4
            Left            =   5730
            TabIndex        =   14
            ToolTipText     =   "Browse Skin directory"
            Top             =   300
            Width           =   465
            _extentx        =   820
            _extenty        =   556
            caption         =   ""
            font            =   "Options.frx":96FB
            picture         =   "Options.frx":9723
            mousehoverpicture=   "Options.frx":A135
            pictureoffsetx  =   120
            pictureoffsety  =   -20
            customstyle     =   0
         End
         Begin MMPlayerXProject.vkListBox Listaskins 
            Height          =   1545
            Left            =   120
            TabIndex        =   28
            Top             =   900
            Width           =   6045
            _extentx        =   10663
            _extenty        =   2725
            backcolor       =   12632256
            bordercolor     =   0
            multiselect     =   0   'False
            sorted          =   0
            font            =   "Options.frx":AB47
            selcolor        =   4210752
            borderselcolor  =   8421504
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Browse for Skin Directory......"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   64
            Top             =   60
            Width           =   2580
         End
         Begin VB.Label lblskin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Skin: DEFAULT"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   660
            Width           =   1995
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   3915
         Index           =   2
         Left            =   7590
         TabIndex        =   19
         Top             =   0
         Width           =   6300
         _extentx        =   11113
         _extenty        =   6906
         font            =   "Options.frx":AB6F
         showtitle       =   0   'False
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin MMPlayerXProject.vkCheck CheckSetting 
            CausesValidation=   0   'False
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   48
            Top             =   360
            Width           =   2955
            _extentx        =   5212
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Automatically play on startup"
            font            =   "Options.frx":AB97
         End
         Begin MMPlayerXProject.vkCheck CheckSetting 
            Height          =   225
            Index           =   6
            Left            =   330
            TabIndex        =   27
            Top             =   2520
            Width           =   4035
            _extentx        =   7117
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "View Player menus in task bar"
            font            =   "Options.frx":ABBF
         End
         Begin MMPlayerXProject.vkCheck CheckSetting 
            Height          =   225
            Index           =   5
            Left            =   330
            TabIndex        =   26
            Top             =   2160
            Width           =   5445
            _extentx        =   9604
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "AutoEnable DSP FX on loading"
            font            =   "Options.frx":ABE7
         End
         Begin MMPlayerXProject.vkCheck CheckSetting 
            Height          =   225
            Index           =   4
            Left            =   330
            TabIndex        =   25
            Top             =   1800
            Width           =   2445
            _extentx        =   4313
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Always on Top"
            font            =   "Options.frx":AC0F
         End
         Begin MMPlayerXProject.vkCheck CheckSetting 
            Height          =   225
            Index           =   3
            Left            =   330
            TabIndex        =   24
            Top             =   1410
            Width           =   2445
            _extentx        =   4313
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Show Splash Screen"
            font            =   "Options.frx":AC37
         End
         Begin MMPlayerXProject.vkCheck CheckSetting 
            Height          =   225
            Index           =   2
            Left            =   330
            TabIndex        =   23
            Top             =   1050
            Width           =   4665
            _extentx        =   8229
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Enable Right Click menu on Directories in Explorer"
            font            =   "Options.frx":AC5F
         End
         Begin MMPlayerXProject.vkCheck CheckSetting 
            Height          =   225
            Index           =   1
            Left            =   330
            TabIndex        =   22
            Top             =   720
            Width           =   2445
            _extentx        =   4313
            _extenty        =   397
            backcolor       =   14737632
            backstyle       =   0
            caption         =   "Allow multiple Instances"
            font            =   "Options.frx":AC87
         End
         Begin VB.Label lblsettings 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Note that  some changes in settings may need restart of player for  their effect to take place"
            Height          =   390
            Index           =   16
            Left            =   300
            TabIndex        =   71
            Top             =   2970
            Width           =   5955
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.FileListBox fileBmps 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Hidden          =   -1  'True
      Left            =   -1260
      Pattern         =   "*.bmp"
      System          =   -1  'True
      TabIndex        =   0
      Top             =   4710
      Visible         =   0   'False
      Width           =   4710
   End
   Begin MMPlayerXProject.Button btnExit 
      Height          =   120
      Left            =   3120
      TabIndex        =   50
      Top             =   0
      Width           =   225
      _extentx        =   397
      _extenty        =   212
      style           =   1
      buttoncolor     =   12632256
      mousepointer    =   99
   End
   Begin VB.PictureBox PicInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   -1860
      ScaleHeight     =   311
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   165
      TabIndex        =   51
      Top             =   0
      Width           =   2475
      Begin MMPlayerXProject.vkCommand CMDV 
         Height          =   945
         Index           =   4
         Left            =   2070
         TabIndex        =   52
         ToolTipText     =   "System Configuartion"
         Top             =   0
         Width           =   405
         _extentx        =   714
         _extenty        =   2196
         backgradient    =   1
         caption         =   ""
         font            =   "Options.frx":ACAF
         breakcorner     =   0   'False
         picture         =   "Options.frx":ACD7
         mousehoverpicture=   "Options.frx":B6E9
         pictureoffsetx  =   100
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand CMDV 
         Height          =   945
         Index           =   3
         Left            =   2070
         TabIndex        =   53
         ToolTipText     =   "Player's Configuration 2"
         Top             =   930
         Width           =   405
         _extentx        =   714
         _extenty        =   1667
         backgradient    =   1
         caption         =   ""
         font            =   "Options.frx":C0FB
         breakcorner     =   0   'False
         picture         =   "Options.frx":C123
         mousehoverpicture=   "Options.frx":CB35
         pictureoffsetx  =   123
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand CMDV 
         Height          =   945
         Index           =   2
         Left            =   2070
         TabIndex        =   54
         ToolTipText     =   "Player's Configuration 1"
         Top             =   1860
         Width           =   405
         _extentx        =   714
         _extenty        =   1667
         backgradient    =   1
         caption         =   ""
         font            =   "Options.frx":D547
         breakcorner     =   0   'False
         picture         =   "Options.frx":D56F
         mousehoverpicture=   "Options.frx":DF81
         pictureoffsetx  =   123
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand CMDV 
         Height          =   945
         Index           =   1
         Left            =   2070
         TabIndex        =   55
         ToolTipText     =   "Skin Configuration"
         Top             =   2790
         Width           =   405
         _extentx        =   714
         _extenty        =   1667
         backgradient    =   1
         caption         =   ""
         font            =   "Options.frx":E993
         breakcorner     =   0   'False
         picture         =   "Options.frx":E9BB
         mousehoverpicture=   "Options.frx":F3CD
         pictureoffsetx  =   123
         customstyle     =   0
      End
      Begin MMPlayerXProject.vkCommand CMDV 
         Height          =   945
         Index           =   0
         Left            =   2070
         TabIndex        =   56
         ToolTipText     =   "Playlist Configuration"
         Top             =   3720
         Width           =   405
         _extentx        =   714
         _extenty        =   1667
         backgradient    =   1
         caption         =   ""
         font            =   "Options.frx":FDDF
         breakcorner     =   0   'False
         picture         =   "Options.frx":FE07
         mousehoverpicture=   "Options.frx":10819
         pictureoffsetx  =   123
         customstyle     =   0
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cWindows As New cWindowSkin
Dim cAjustarDesk As New clsDockingHandler

Dim ListScroll As New vkPrivateScroll
Private BackColor1 As Long
Private BackColor2 As Long
Private backColorSelect As Long
Public activeTab As Long


Private Sub btnExit_Click()
    frmPopUp.mnuPreferences_Click
End Sub

Private Sub CheckSetting_Change(Index As Integer, Value As CheckBoxConstants)
    Select Case Index
    Case 0:
        bPlayStarting = CheckSetting(Index).Value
    Case 1:
        OpcionesMusic.Instancias = CheckSetting(Index).Value
    Case 2:
        EnableRightClickmenu CheckSetting(Index).Value
    Case 3:
        OpcionesMusic.Splash = CheckSetting(Index).Value
    Case 4:
        OpcionesMusic.SiempreTop = CheckSetting(Index).Value
        Always_on_Top
    Case 5:
        bPlayDFXStarting = CheckSetting(Index).Value
    Case 6:
        OpcionesMusic.SysMenu = CheckSetting(Index).Value
    End Select
End Sub

Private Sub chkUsefile_Change(Value As CheckBoxConstants)
    bLoadRegionFile = chkUsefile.Value
End Sub

Private Sub CMD_Click(Index As Integer)
'There are 4 vkCommand buttons available in each tab
'index=0 (default),1=Cancel,2=OK,3=Apply
'index=4 for browsing path for skin directory
    Select Case Index
        Dim k
        Dim i As Long

    Case 0    'LOAD DEFAULTS
        Load_Defaults
        frmMain.Show_Message "Default Settings Loaded"

    Case 1    'CANCEL
        frmPopUp.mnuPreferences_Click

    Case 2    'OK
        CMD_Click (4)
        frmPopUp.mnuPreferences_Click

    Case 3    'APPLY
        '*************Playlist format******
        sFormatScroll = UCase(Trim(txtDisplay(1).Text))
        If sFormatScroll = "" Then sFormatScroll = "%S - %A (%T)"
        'sFormatPlayList text is changed then prompt for list update
        If UCase(Trim(txtDisplay(0).Text)) <> UCase(sFormatPlayList) Then
            If sFormatPlayList = "" Then sFormatPlayList = "%A - %S"
            sFormatPlayList = UCase((txtDisplay(0).Text))
            If frmPLST.clist.ItemCount > 0 Then k = MsgBox("Do you want to update playlist in new format? It may take few seconds...", vbYesNo, "Neo player")
            'k=1 means ok k=2 means Cancel
            If k = 6 Then Call frmPLST.UpdatePlaylist    'k<>0 shows yes
        End If

        '*************Register files******
        For i = 1 To vkListType.ListCount
            Call RegisterFiles(LCase(Left(vkListType.Item(i).Text, 3)), vkListType.Item(i).Checked)
        Next

        '*************Change Skin*********
        Apply_Skin

    Case 4    'LOAD SKIN PATH
        On Error Resume Next
        Dim strSkinPath As String
        'strSkinPath = Explorador_Para_Directorios(Me.hWnd, LineLanguage(76))
        If Trim(strSkinPath) = "" Then Exit Sub
        If Right(strSkinPath, 1) <> "\" Then strSkinPath = strSkinPath & "\"
        tAppConfig.AppConfig = strSkinPath
        txtAppConfig.Text = strSkinPath
        Search_Skins_Languages
        Load_Skins_Menu ""

    End Select


End Sub
Private Sub CMDV_Click(Index As Integer)
    PicInfo.cls
    If Index = 1 Then
        'for screen shot ...currently disabled
        'BitBlt PicInfo.hDc, 10, 6, picscreenshot.Width, picscreenshot.Height / 2, picscreenshot.hDc, 0, 0, &H8800C6
        'BitBlt PicInfo.hDc, 10, 6, picscreenshot.Width, picscreenshot.Height / 2, picscreenshot.hDc, 0, picscreenshot.Height / 2, &H660046
    End If
    activeTab = Index
    Dim i
    For i = 0 To 4
        CMDV(i).BackColor1 = BackColor1
        CMDV(i).BackColor2 = BackColor2
    Next
    CMDV(Index).BackColor1 = backColorSelect
    CMDV(Index).BackColor2 = backColorSelect
    frameBack(Index).ZOrder vbBringToFront
End Sub

Private Sub Form_Load()
    LoadSkin
    Const flag As Long = SWP_NOMOVE Or SWP_NOSIZE
    ' boolOptionsShow = True this should be decided by other form event
    activeTab = 0
    LoadConf
    boolOptionsLoaded = True
    If OpcionesMusic.SiempreTop = True Then
        SetWindowPos frmOptions.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flag
    Else
        SetWindowPos frmOptions.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flag
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, 161, 2, 0
End Sub

Private Sub LoadConfig()
    Dim Labelforecolor
    'Debug.Print tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini"
    PicInfo.Left = frameContainer.Left - PicInfo.Width + 1
    PicInfo.Top = frameContainer.Top
    BackColor1 = cRead_INI("BUTTON", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    BackColor2 = cRead_INI("BUTTON", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    backColorSelect = cRead_INI("BUTTON", "BackcolorSelect", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    frameContainer.BackColor1 = cRead_INI("CONTAINER", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    frameContainer.BackColor2 = cRead_INI("CONTAINER", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    Labelforecolor = cRead_INI("LABEL", "forecolor", RGB(0, 255, 0), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")

    Dim i
    For i = 0 To 4
        CMD(i).BackColor1 = BackColor1
        CMD(i).BackColor2 = BackColor2
        CMD(i).ForeColor = cRead_INI("BUTTON", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
        CMD(i).BorderColor = cRead_INI("BUTTON", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
        CMD(i).Top = 4080
        CMDV(i).BackColor1 = BackColor1
        CMDV(i).BackColor2 = BackColor2
        CMDV(i).ForeColor = cRead_INI("BUTTON", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
        CMDV(i).BorderColor = cRead_INI("BUTTON", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    Next
    CMD(4).Top = 300    ' button kept in different container = Browse for directory button
    CMD(4).BorderColor = RGB(20, 10, 20)    '

    For i = 0 To 4
        frameBack(i).BackColor2 = cRead_INI("CONTAINER", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
        frameBack(i).BackColor1 = cRead_INI("CONTAINER", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
        frameBack(i).Left = 57
        frameBack(i).Top = 59
        frameBack(i).Width = 6260
    Next
    For i = 0 To 6
        CheckSetting(i).BackColor = frameBack(0).BackColor2
    Next
    chkUsefile.BackColor = frameBack(1).BackColor2
    For i = 0 To 1
        optScrollType(i).BackColor = frameBack(0).BackColor2
    Next

    For i = 0 To 2
        optPlstCase(i).BackColor = frameBack(0).BackColor2
    Next

    For i = 0 To 6
        vkCheckIcon(i).BackColor = frameBack(4).BackColor2
    Next

    For i = 0 To 16
        lblsettings(i).ForeColor = Labelforecolor
        If i <= 2 Then optPlstCase(i).ForeColor = Labelforecolor
        If i <= 6 Then CheckSetting(i).ForeColor = Labelforecolor
        If i <= 1 Then optScrollType(i).ForeColor = Labelforecolor
        If i <= 6 Then vkCheckIcon(i).ForeColor = Labelforecolor
    Next
    ' optScrollType(0).ForeColor = cRead_INI("CHECKBOX", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    ' optScrollType(1).ForeColor = cRead_INI("CHECKBOX", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")



    picscreenshot.BackColor = cRead_INI("TEXT", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    ListScroll.BackColor = cRead_INI("LIST", "ScrollbackColor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    ListScroll.FrontColor = cRead_INI("LIST", "ScrollfrontColor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    ListScroll.BorderColor = cRead_INI("LIST", "ScrollborderColor", RGB(0, 0, 0), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    ListScroll.ArrowColor = cRead_INI("LIST", "ScrollarrowColor", RGB(0, 0, 0), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")

    txtDisplay(1).ForeColor = cRead_INI("TEXT", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtDisplay(0).ForeColor = cRead_INI("TEXT", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtFormat.ForeColor = cRead_INI("TEXT", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtAppConfig.ForeColor = cRead_INI("TEXT", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")

    'chkscrolltype
    Listaskins.ForeColor = cRead_INI("LIST", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    Listaskins.BorderColor = cRead_INI("LIST", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    Listaskins.BackColor = cRead_INI("LIST", "backcolor", RGB(70, 100, 50), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    Listaskins.SelColor = cRead_INI("LIST", "selcolor", RGB(70, 100, 50), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    Listaskins.BackColor = cRead_INI("LIST", "backcolor", RGB(70, 100, 50), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtAppConfig.BackColor = Listaskins.BackColor

    vkListType.BackColor = Listaskins.BackColor
    vkListType.BorderColor = Listaskins.BorderColor
    vkListType.ForeColor = Listaskins.ForeColor
    vkListType.SelColor = Listaskins.SelColor
    Listaskins.VScroll = ListScroll
    vkListType.VScroll = Listaskins.VScroll

    txtSkininfo.ForeColor = cRead_INI("TEXT", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtSkininfo.BackColor = cRead_INI("TEXT", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtDisplay(0).BackColor = cRead_INI("TEXT", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtFormat.BackColor = cRead_INI("TEXT", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtDisplay(1).BackColor = cRead_INI("TEXT", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    ' vkScrollUpDown.FrontColor = cRead_INI("TEXT", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    txtAppConfig.BackColor = cRead_INI("TEXT", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")

End Sub



Private Sub Apply_Skin()
'On Error Resume Next
    Dim Skin As String
    Dim k
    If Listaskins.ListIndex < 0 Or Listaskins.ListCount = 0 Then Exit Sub

    Skin = Trim(Listaskins.Item(Listaskins.ListIndex).Text)
    'If selected skin is other than current skin then prompt for changing skin
    If Skin <> tAppConfig.Skin Then
        k = MsgBox("Do u want to change the skin from " & tAppConfig.Skin & " to " & Skin & ". It may take few seconds please keep patience...", vbOKCancel, "Neo player")
        'k=1 means ok k=2 means Cancel
        If k = 1 Then Call frmPopUp.mnuSkinsAdd_Click(Listaskins.ListIndex): lblskin.Caption = "Current Skin: " & tAppConfig.Skin
    End If

End Sub

Public Sub LoadSkin()
    Dim k
    On Error Resume Next
    Me.Height = Read_INI("FORM", "formheight", 6020, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    Me.Width = Read_INI("FORM", "formwidth", 8888, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")

    Set cWindows.FormularioPadre = Me
    Set cAjustarDesk.ParentForm = Me
    cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
    cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\", True
    frameContainer.Top = Read_INI("CONTAINER", "top", 12, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    frameContainer.Left = Read_INI("CONTAINER", "left", 15, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    frameContainer.Width = 426    'same for all skins
    frameContainer.Height = 313

    'set exit buuton at place
    k = Read_Config_Button(btnExit, "Configuration", "exitButton", "0,0,10,10", tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\CONFIG\config.ini")
    Set btnExit.PictureNormal = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\Exitnormal.bmp")
    Set btnExit.PictureOver = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\Exitover.bmp")
    Set btnExit.PictureDown = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\ExitDown.bmp")
    Set btnExit.PictureNormal = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\Exitnormal.bmp")

    LoadConfig
    ' Set ScrollText.PictureText = frmMain.ScrollText(5).PictureText
    ' ScrollText.Height = frmMain.ScrollText(1).Height
    '  ScrollText.Left = 30
    '  ScrollText.Top = 360
    PicInfo.BackColor = frmPLST.picList.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    boolOptionsShow = False
End Sub

Sub Search_Skins_Languages()
    Dim miNombre As String
    Dim sPathskin As String
    'Function Dir([PathName], [Attributes As VbFileAttribute = vbNormal]) As String
    'Returns the name of a matching file, directory, or folder

    On Error Resume Next
    'search skins in musicmp3/skins only directories
    fileBmps.Pattern = "*.bmp"
    Listaskins.Clear
    miNombre = Dir(tAppConfig.AppConfig & "Skins\", vbDirectory)  '//Skins root diretory
    sPathskin = tAppConfig.AppConfig & "Skins\"
    Dim i
    i = 0
    Do While miNombre <> ""
        If miNombre <> "." And miNombre <> ".." Then
            'check for each directry inside "skin" directiry
            If (GetAttr(sPathskin & miNombre) And vbDirectory) = vbDirectory Then
                fileBmps.Path = tAppConfig.AppConfig & "Skins\" & miNombre
                '// check if any bmp/jpeg file is there
                If fileBmps.ListCount > 0 Then
                    Listaskins.AddItem miNombre
                    '//we have found the skin
                    If LCase(Trim(miNombre)) = LCase(Trim(tAppConfig.Skin)) Then
                        Listaskins.Selected(i + 1) = True
                        Listaskins_ItemClick Listaskins.Item(i + 1)
                    End If
                    i = i + 1
                End If
            End If
        End If
        miNombre = Dir    ' check for next directory
    Loop
    'Listaskins.Selected(
End Sub

Private Sub frameBack_MouseDown(Index As Integer, Button As MouseButtonConstants, Shift As Integer, Control As Integer, X As Long, Y As Long)
'ReleaseCapture
'  SendMessage frameBack(Index).hWnd, 161, 2, 0
End Sub

Private Sub Listaskins_ItemClick(Item As vkListItem)
    If Item Is Nothing Then Exit Sub    ' error comes when clicked in clear region so 'nothing' item is returned

    Dim strInfo As String, strSkinTemp As String
    strSkinTemp = tAppConfig.AppConfig & "Skins\" & Item.Text & "\skin.ini"
    strInfo = Read_INI("Info", "AuthorName", "", , , strSkinTemp)
    txtSkininfo.Text = "Author: " & strInfo
    strInfo = Read_INI("Info", "Email", "", , , strSkinTemp)
    txtSkininfo.Text = txtSkininfo.Text & vbCrLf & "E-mail: " & strInfo
    strInfo = Read_INI("Info", "Comments", "", , , strSkinTemp)
    txtSkininfo.Text = txtSkininfo.Text & vbCrLf & "Comments: " & strInfo
    's = tAppConfig.AppConfig & "Skins\" & Item.Text & "\"
    'picscreenshot.Picture = LoadPicture()
    'picscreenshot.Picture = LoadPicture(s & "Screenshot.bmp")
    'PicInfo.Cls
    'BitBlt PicInfo.hDc, 10, 6, picscreenshot.Width, picscreenshot.Height / 2, picscreenshot.hDc, 0, 0, &H8800C6
    'BitBlt PicInfo.hDc, 10, 6, picscreenshot.Width, picscreenshot.Height / 2, picscreenshot.hDc, 0, picscreenshot.Height / 2, &H660046
End Sub

Private Sub Listaskins_ItemDblClick(Item As vkListItem)
    On Error GoTo hell
    If Item.Text <> "" Then Apply_Skin
hell:
End Sub


Public Sub AddfileTypes()
    vkListType.AddItem "MP1 Files"
    vkListType.AddItem "MP2 Files"
    vkListType.AddItem "MP3 Files"
    vkListType.AddItem "Wav Files"
    vkListType.AddItem "Ogg Files"
    vkListType.AddItem "Wma Files"
    vkListType.AddItem "Npl Files"
    vkListType.AddItem "M3u Files"
    vkListType.AddItem "Pls Files"
    ' vkListType.AddItem "Audio CD Files"

    '  vkListType.AddItem "Amr Files"
    ' vkListType.AddItem "MPEG Files"
    'vkListType.AddItem "MPEG2 Files"
    'vkListType.AddItem "MPEG3 Files"
    'vkListType.AddItem "MIDI Files"
    'vkListType.AddItem "Aac Files"
End Sub


Private Sub OptPlstCase_Change(Index As Integer, Value As CheckBoxConstants)
    iOptPlstCase = Index
    frmPLST.ReinitializeList
End Sub

Private Sub optscrolltype_Change(Index As Integer, Value As CheckBoxConstants)
'If bformloading = True Then Exit Sub
    Select Case Index
    Case 0  'Rolling
        frmMain.ScrollText(1).ScrollType = Rolling
        frmMain.ScrollText(5).ScrollType = Rolling
        iScrollType = 0
    Case 1  'Zig ZAg
        frmMain.ScrollText(1).ScrollType = ZigZag
        frmMain.ScrollText(5).ScrollType = ZigZag
        iScrollType = 1
    End Select
End Sub

Private Sub sldSettings_Change(Index As Integer, Value As Long)
'On Error GoTo hell
    Select Case Index

    Case 0    '// tracks
        iCrossfadeTrack = sldSettings(Index).Value
        frmMain.Show_Message "T-Crossfade: " + str(sldSettings(Index).Value) + " ms"
        lblsettings(Index).Caption = "Crossfade between tracks (in ms):" + str(sldSettings(Index).Value) + " ms"
    Case 1    '// stop
        iCrossfadeStop = sldSettings(Index).Value
        frmMain.Show_Message "S-Crossfade: " + str(sldSettings(Index).Value) + " ms"
        lblsettings(Index).Caption = "Crossfade in Stop/Pause(in ms):" + str(sldSettings(Index).Value) + " ms"

    Case 2
        '// Ajust percentage
        sldSettings(Index).ToolTipText = (sldSettings(Index).Value * 100) / 100 & "%"
        Make_Transparent frmMain.hwnd, sldSettings(Index).Value
        Make_Transparent frmPLST.hwnd, sldSettings(Index).Value
        OpcionesMusic.Alpha = sldSettings(Index).Value
        Dim i
        For i = 0 To 9
            frmPopUp.mnuAlpha(i).Checked = False
        Next i

        frmPopUp.mnuAlphaPer.Caption = Trim(LineLanguage(34)) & " [ " & sldSettings(Index).Value & "% ]"
        frmPopUp.mnuAlphaPer.Checked = True
        frmMain.Show_Message "Opacity: " + str(sldSettings(Index).Value) + "%"
        lblsettings(Index).Caption = "Opacity(in %): " + str(sldSettings(Index).Value) + "%"

    Case 3
        frmMain.ScrollText(1).ScrollVelocity = sldSettings(Index).Value
        frmMain.ScrollText(5).ScrollVelocity = sldSettings(Index).Value
        iScrollVel = sldSettings(Index).Value
        frmMain.Timer_Wait.Enabled = False
        frmMain.Show_Message "Scroll vel: " + str(sldSettings(Index).Value)
        lblsettings(3).Caption = "Scrollling text Velocity (in ms):" + str(sldSettings(3).Value)

    Case 4
        iSpectrum_refreshRate = sldSettings(Index).Value
        frmMain.Timer_Player.Interval = sldSettings(Index).Value
        frmMain.Show_Message "Spec.Rate: " + str(sldSettings(Index).Value)
        lblsettings(4).Caption = "Spectrum refresh interval (in ms):" + str(sldSettings(4).Value)
    Case 5
        iVolScrollChange = sldSettings(Index).Value
        frmMain.Show_Message "VolScroll: " + str(sldSettings(Index).Value)
        lblsettings(5).Caption = "Seek Volume in 1 mouse scroll (in %):" + str(sldSettings(5).Value)

    Case 6
        iPosScrollChange = sldSettings(Index).Value
        frmMain.Show_Message "PosScroll: " + str(sldSettings(Index).Value)
        lblsettings(6).Caption = "Seek track in 1 mouse scroll (in s):" + str(sldSettings(6).Value)

    End Select

hell:
End Sub




Public Sub LoadConf()
    txtAppConfig.Text = tAppConfig.AppConfig
    If bLoadRegionFile = True Then chkUsefile.Value = vbChecked
    Listaskins.Clear
    Search_Skins_Languages
    lblskin.Caption = "Current Skin: " & tAppConfig.Skin
    AddfileTypes    ' add entries in filetype listbox

    'initialize sliders
    sldSettings(0).Value = iCrossfadeTrack    '100 default
    sldSettings(1).Value = iCrossfadeStop    '100 default
    sldSettings(2).Value = OpcionesMusic.Alpha  '100// alpha slider
    sldSettings(3).Value = iScrollVel  ' 130 default
    sldSettings(4).Value = iSpectrum_refreshRate    '50
    sldSettings(5).Value = iVolScrollChange    '5 default
    sldSettings(6).Value = iPosScrollChange    '5 default
    optPlstCase(iOptPlstCase).Value = vbChecked
    txtDisplay(0).Text = sFormatPlayList

    On Error Resume Next
    If OpcionesMusic.TaskBar = True Then vkCheckIcon(6).Value = vbChecked

    If OpcionesMusic.SysTray = True Then vkCheckIcon(0).Value = vbChecked

    '// Player icons
    If PlayerTrayIcon.Previous = True Then vkCheckIcon(1).Value = vbChecked
    If PlayerTrayIcon.Play = True Then vkCheckIcon(2).Value = vbChecked
    If PlayerTrayIcon.Pause = True Then vkCheckIcon(3).Value = vbChecked
    If PlayerTrayIcon.Stop = True Then vkCheckIcon(4).Value = vbChecked
    If PlayerTrayIcon.Next = True Then vkCheckIcon(5).Value = vbChecked
    sldSettings(4).Value = iSpectrum_refreshRate

    '// scroll caption
    txtDisplay(1).Text = sFormatScroll
    txtDisplay(0).Text = sFormatPlayList

    optScrollType(iScrollType).Value = vbChecked

    txtAppConfig.Text = tAppConfig.AppConfig

    If bLoadRegionFile = True Then chkUsefile.Value = vbChecked

    If bPlayStarting = True Then CheckSetting(0).Value = vbChecked
    If OpcionesMusic.Instancias = True Then CheckSetting(1).Value = vbChecked
    If OpcionesMusic.Directorio = True Then CheckSetting(2).Value = vbChecked
    If OpcionesMusic.Splash = True Then CheckSetting(3).Value = vbChecked
    If OpcionesMusic.SiempreTop = True Then CheckSetting(4).Value = vbChecked
    If bPlayDFXStarting = True Then CheckSetting(5).Value = vbChecked
    If OpcionesMusic.SysMenu = True Then CheckSetting(6).Value = vbChecked

End Sub

Public Sub EnableRightClickmenu(bEnable As Boolean)
    On Error Resume Next
    Dim lngRootKey As Long
    Dim RutaExe As String
    lngRootKey = HKEY_CLASSES_ROOT

    '  If bLoading = True Then Exit Sub
    '+--------------------------------------------------------------------------------+
    '|procedure for access to the registry when           |
    '|rightclicking in a folder or driver displayed the text 'Search Music Mp3 Player X'
    '|and run the application with the parameters sent in this case where we
    '|right click
    '|keys are:                                                                    |
    '| --> HKEY_CLASSES_ROOT\Directory\Shell\ 'Texto del Menu'                           |
    '| --> HKEY_CLASSES_ROOT\Directory\Shell\ 'Texto del Menu' \command                  |
    '|                                  with a key to the path of the application and     |
    '|                                  command                                         |
    '+--------------------------------------------------------------------------------+
    If bEnable Then
        OpcionesMusic.Directorio = True
        '// obtain the correct string to put in the record
        RutaExe = tAppConfig.AppPath & App.EXEName & ".exe %1"
        'Verifika if key exists
        If Not regDoes_Key_Exist(lngRootKey, "Directory\shell\Play with NeoMp3 Player") Then
            regCreate_A_Key lngRootKey, "Directory\shell\Play with NeoMp3 Player"
            regCreate_A_Key lngRootKey, "Directory\shell\Play with NeoMp3 Player\command"
            RutaExe = tAppConfig.AppPath & App.EXEName & ".exe RUND %1"
            regCreate_Key_Value lngRootKey, "Directory\shell\Play with NeoMp3 Player\command", "", RutaExe

            regCreate_A_Key lngRootKey, "Directory\shell\Enqueue in NeoMp3 Player"
            regCreate_A_Key lngRootKey, "Directory\shell\Enqueue in NeoMp3 Player\command"
            RutaExe = tAppConfig.AppPath & App.EXEName & ".exe ADDD %1"
            regCreate_Key_Value lngRootKey, "Directory\shell\Enqueue in NeoMp3 Player\command", "", RutaExe

        End If

        If Not regDoes_Key_Exist(lngRootKey, "Drive\shell\Search in NeoMp3 Player") Then
            RutaExe = tAppConfig.AppPath & App.EXEName & ".exe FIND %1"
            regCreate_A_Key lngRootKey, "Drive\shell\Search in NeoMp3 Player"
            regCreate_A_Key lngRootKey, "Drive\shell\Search in NeoMp3 Player\command"
            regCreate_Key_Value lngRootKey, "Drive\shell\Search in NeoMp3 Player\command", "", RutaExe
        End If
    Else
        OpcionesMusic.Directorio = False
        regDelete_A_Key lngRootKey, "Directory\shell\Play with NeoMp3 Player", "command"
        regDelete_A_Key lngRootKey, "Directory\shell", "Play with NeoMp3 Player"
        regDelete_A_Key lngRootKey, "Directory\shell\Enqueue in NeoMp3 Player", "command"
        regDelete_A_Key lngRootKey, "Directory\shell", "Enqueue in NeoMp3 Player"
        regDelete_A_Key lngRootKey, "Drive\shell\Search in NeoMp3 Player", "command"
        regDelete_A_Key lngRootKey, "Drive\shell", "Search in NeoMp3 Player"
    End If
End Sub
Public Sub RegisterFiles(sfileExt As String, bRegister As Boolean)
    Select Case sfileExt
    Case "mp3"
        If bRegister = True Then
            Call RegisterType(".mp3", "NeoMP3.File", "AUDIO FILES", "Neo Media File", 0)
        Else
            Call DeleteType(".mp3", "NeoMP3.File")
        End If
    Case "wav"
        If bRegister = True Then
            Call RegisterType(".wav", "Neo wave File", "AUDIO FILES", "Neo Media File", 0)
        Else
            Call DeleteType(".wav", "Neo wave File")
        End If
    Case "ogg"
        If bRegister = True Then
            Call RegisterType(".ogg", "Neo ogg File", "AUDIO FILES", "Neo ogg vorbis File", 0)
        Else
            Call DeleteType(".ogg", "Neo ogg File")
        End If
    Case "wma"
        If bRegister = True Then
            Call RegisterType(".wma", "Neo wma File", "AUDIO FILES", "Neo Media File", 0)
        Else
            Call DeleteType(".wma", "Neo wma File")
        End If

    Case "npl"
        If bRegister = True Then
            Call RegisterType(".npl", "Neo npl File", "PLAYLIST FILES", "Neo playlist File", 1)
        Else
            Call DeleteType(".npl", "Neo npl File")
        End If
    Case "m3u"
        If bRegister = True Then
            Call RegisterType(".m3u", "Neo m3u File", "PLAYLIST FILES", "Neo playlist File", 1)
        Else
            Call DeleteType(".m3u", "Neo m3u File")
        End If
    Case "pls"
        If bRegister = True Then
            Call RegisterType(".pls", "Neo pls File", "PLAYLIST FILES", "Neo playlist File", 1)
        Else
            Call DeleteType(".pls", "Neo pls File")
        End If
    End Select



End Sub




Public Sub Load_Defaults()

    Dim i As Long
    '*************Default Skin******
    If Listaskins.ListCount > 0 Then
        Listaskins.UnSelectAll: Listaskins.TopIndex = 1: Listaskins.Selected(1) = True
        Listaskins_ItemClick Listaskins.Item(1)
        Listaskins.Refresh
    End If

    chkUsefile.Value = vbChecked
    '*************Default Playlist******
    sFormatScroll = "%S - %A (%T)"
    txtDisplay(1).Text = sFormatScroll
    sFormatPlayList = "%A - %S"
    txtDisplay(0).Text = sFormatPlayList
    optPlstCase(0).Value = vbChecked    'set normal mode as default
    optScrollType(0).Value = vbChecked    'set zig zag scrolling as default

    '*************System Tray******
    For i = 1 To 5
        vkCheckIcon(i).Value = vbUnchecked
    Next
    vkCheckIcon(0).Value = vbChecked
    vkCheckIcon(6).Value = vbChecked

    vkListType.UnCheckAll
    vkListType.Checked(3) = True    'mp3
    vkListType.Checked(7) = True    'npl
    vkListType.Refresh
    '*************Configuration******
    For i = 0 To 6
        If i <> 1 And i <> 5 Then CheckSetting(i).Value = vbChecked
    Next
    CheckSetting(1).Value = vbUnchecked    'Allow multiple instances
    CheckSetting(5).Value = vbUnchecked    'auto enable DFX

    '*************Slider settings******
    sldSettings(0).Value = 100    ' iCrossfadeTrack '100 default
    sldSettings(1).Value = 100    'iCrossfadeStop '100 default
    sldSettings(2).Value = 100    'OpcionesMusic.Alpha  '100// alpha slider
    sldSettings(3).Value = 130    'iScrollVel  ' 130 default
    sldSettings(4).Value = 50    'iSpectrum_refreshRate '50
    sldSettings(5).Value = 5    'iVolScrollChange '5 default
    sldSettings(6).Value = 5    'iPosScrollChange '5 default

End Sub

Private Sub vkCheckIcon_Change(Index As Integer, Value As CheckBoxConstants)
' If bFormloading = True Then Exit Sub

    If Index < 6 Then
        If vkCheckIcon(Index).Value = vbChecked Then
            frmPopUp.vkSysTrayIcon(Index).AddToTray (Index): OpcionesMusic.SysTray = True
        Else
            frmPopUp.vkSysTrayIcon(Index).RemoveFromTray (Index): OpcionesMusic.SysTray = False
        End If
    End If

    Select Case Index
    Case 0  '// Show in task bar
        OpcionesMusic.TaskBar = vkCheckIcon(0).Value
        frmPopUp.Visible = vkCheckIcon(0).Value
    Case 1
        PlayerTrayIcon.Previous = vkCheckIcon(1).Value
    Case 2
        PlayerTrayIcon.Play = vkCheckIcon(2).Value
    Case 3
        PlayerTrayIcon.Pause = vkCheckIcon(3).Value
    Case 4
        PlayerTrayIcon.Stop = vkCheckIcon(4).Value
    Case 5
        PlayerTrayIcon.Next = vkCheckIcon(5).Value
    End Select
End Sub





