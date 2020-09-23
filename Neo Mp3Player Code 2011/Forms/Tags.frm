VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTags 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   8430
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11070
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Tags.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   562
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   738
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdundo 
      Caption         =   "Command1"
      Height          =   555
      Left            =   4830
      TabIndex        =   4
      Top             =   7410
      Width           =   1275
   End
   Begin ComctlLib.ListView listRef 
      Height          =   645
      Left            =   1530
      TabIndex        =   1
      Top             =   7650
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   1138
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   17
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "FILE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "TRACK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "TITLE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ARTIST"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ALBUM"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "YEAR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "GENRE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "COMMENTS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "COMPOSER"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ORG ARTIST"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "LINK"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ENCODER"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   12
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "IMAGE"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   13
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "LYRICS"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   14
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "FILEPATH"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   15
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Edit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   16
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Row"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   18930
      Index           =   2
      Left            =   2670
      TabIndex        =   0
      Top             =   12180
      Width           =   18525
   End
   Begin MMPlayerXProject.vkFrame frameBack 
      Height          =   855
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   7530
      Visible         =   0   'False
      Width           =   1170
      _extentx        =   2064
      _extenty        =   1508
      font            =   "Tags.frx":000C
      showtitle       =   0
      titlecolor1     =   12632256
      bordercolor     =   0
      roundangle      =   0
   End
   Begin MMPlayerXProject.Button btnExit 
      Height          =   120
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   225
      _extentx        =   397
      _extenty        =   212
      style           =   1
      buttoncolor     =   12632256
      mousepointer    =   99
   End
   Begin MMPlayerXProject.vkFrame frameBackmain 
      Height          =   6030
      Left            =   330
      TabIndex        =   5
      Top             =   450
      Width           =   8970
      _extentx        =   15822
      _extenty        =   10478
      backcolor1      =   14737632
      font            =   "Tags.frx":0034
      showtitle       =   0
      titlecolor1     =   12632256
      bordercolor     =   0
      roundangle      =   0
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   5205
         Index           =   2
         Left            =   90
         TabIndex        =   65
         Top             =   600
         Width           =   3450
         _extentx        =   6085
         _extenty        =   8229
         backcolor1      =   14737632
         backcolor2      =   14737632
         font            =   "Tags.frx":005C
         showtitle       =   0
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         displaypicture  =   0
         Begin MMPlayerXProject.vkListBox vkFiletags 
            Height          =   4245
            Left            =   60
            TabIndex        =   67
            Top             =   330
            Width           =   3315
            _extentx        =   5847
            _extenty        =   7382
            backcolor       =   14737632
            bordercolor     =   0
            multiselect     =   0
            sorted          =   0
            stylecheckbox   =   -1
            font            =   "Tags.frx":0084
            selcolor        =   4210752
            borderselcolor  =   12632256
            usedefautitemsettings=   -1
            oledropmode     =   1
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   3
            Left            =   720
            TabIndex        =   66
            Top             =   4680
            Width           =   1905
            _extentx        =   3360
            _extenty        =   714
            caption         =   "Select All"
            font            =   "Tags.frx":00AC
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Track List Information:"
            Height          =   195
            Index           =   12
            Left            =   45
            TabIndex        =   68
            Top             =   60
            Width           =   1965
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   4245
         Index           =   4
         Left            =   3630
         TabIndex        =   57
         Top             =   960
         Width           =   5160
         _extentx        =   8943
         _extenty        =   8070
         font            =   "Tags.frx":00D4
         showtitle       =   0
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   375
            Index           =   11
            Left            =   360
            TabIndex        =   72
            Top             =   1290
            Width           =   1395
            _extentx        =   2461
            _extenty        =   661
            caption         =   "Save Art As"
            font            =   "Tags.frx":00FC
            enabled         =   0
            customstyle     =   0
         End
         Begin MMPlayerXProject.vkUpDown vkUpDownArt 
            Height          =   255
            Left            =   4455
            TabIndex        =   71
            Top             =   1660
            Visible         =   0   'False
            Width           =   375
            _extentx        =   661
            _extenty        =   450
            max             =   1
            min             =   1
            value           =   1
            mouseinterval   =   0
            direction       =   0
         End
         Begin VB.PictureBox PicAlbumArt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FBFBFB&
            ForeColor       =   &H80000008&
            Height          =   1575
            Left            =   3060
            ScaleHeight     =   103
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   117
            TabIndex        =   70
            Top             =   360
            Width           =   1785
         End
         Begin MMPlayerXProject.vkCheck chktags 
            CausesValidation=   0   'False
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   60
            Top             =   1710
            Visible         =   0   'False
            Width           =   225
            _extentx        =   397
            _extenty        =   450
            backcolor       =   16777215
            backstyle       =   0
            caption         =   ""
            font            =   "Tags.frx":0124
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   7
            Left            =   360
            TabIndex        =   59
            Top             =   390
            Width           =   1395
            _extentx        =   2461
            _extenty        =   714
            caption         =   "Add Art"
            font            =   "Tags.frx":014C
            enabled         =   0
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   8
            Left            =   360
            TabIndex        =   58
            Top             =   840
            Width           =   1395
            _extentx        =   2461
            _extenty        =   714
            caption         =   "Remove Art"
            font            =   "Tags.frx":0174
            enabled         =   0
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Update Art Image"
            Enabled         =   0   'False
            Height          =   195
            Index           =   11
            Left            =   660
            TabIndex        =   64
            Top             =   1740
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.Shape Shape2 
            Height          =   1755
            Left            =   180
            Top             =   270
            Width           =   4755
         End
         Begin VB.Shape Shape3 
            Height          =   1785
            Left            =   180
            Top             =   2340
            Width           =   4755
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Album Art:"
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   63
            Top             =   60
            Width           =   930
         End
         Begin VB.Label Label 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MPEG Info:"
            Height          =   195
            Index           =   14
            Left            =   195
            TabIndex        =   62
            Top             =   2100
            Width           =   960
         End
         Begin VB.Label lblMPEGInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   1755
            Left            =   240
            TabIndex        =   61
            Top             =   2390
            Width           =   4650
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00E0E0E0&
            BorderWidth     =   2
            X1              =   0
            X2              =   1340
            Y1              =   0
            Y2              =   0
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   4245
         Index           =   5
         Left            =   3630
         TabIndex        =   48
         Top             =   960
         Width           =   5160
         _extentx        =   9102
         _extenty        =   7488
         font            =   "Tags.frx":019C
         showtitle       =   0
         titlecolor1     =   12632256
         bordercolor     =   0
         roundangle      =   0
         Begin MMPlayerXProject.vkCommand vkcmdPlayer 
            Height          =   315
            Index           =   0
            Left            =   1170
            TabIndex        =   56
            Top             =   270
            Width           =   405
            _extentx        =   2461
            _extenty        =   714
            caption         =   "<<"
            font            =   "Tags.frx":01C4
         End
         Begin MMPlayerXProject.vkCommand vkcmdPlayer 
            Height          =   315
            Index           =   1
            Left            =   1650
            TabIndex        =   55
            Top             =   270
            Width           =   405
            _extentx        =   2461
            _extenty        =   714
            caption         =   ">"
            font            =   "Tags.frx":01EC
         End
         Begin MMPlayerXProject.vkCommand vkcmdPlayer 
            Height          =   315
            Index           =   2
            Left            =   2130
            TabIndex        =   54
            Top             =   270
            Width           =   405
            _extentx        =   2461
            _extenty        =   714
            caption         =   "||"
            font            =   "Tags.frx":0214
         End
         Begin MMPlayerXProject.vkCommand vkcmdPlayer 
            Height          =   315
            Index           =   3
            Left            =   2610
            TabIndex        =   53
            Top             =   270
            Width           =   405
            _extentx        =   2461
            _extenty        =   714
            caption         =   "X"
            font            =   "Tags.frx":023C
         End
         Begin MMPlayerXProject.vkCommand vkcmdPlayer 
            Height          =   315
            Index           =   4
            Left            =   3090
            TabIndex        =   52
            Top             =   270
            Width           =   405
            _extentx        =   2461
            _extenty        =   714
            caption         =   ">>"
            font            =   "Tags.frx":0264
         End
         Begin VB.TextBox txtLyrics 
            Appearance      =   0  'Flat
            Height          =   2910
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   51
            Top             =   690
            Width           =   4905
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   375
            Index           =   9
            Left            =   810
            TabIndex        =   50
            Top             =   3690
            Width           =   1395
            _extentx        =   2461
            _extenty        =   661
            caption         =   "Add"
            font            =   "Tags.frx":028C
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   375
            Index           =   10
            Left            =   2670
            TabIndex        =   49
            Top             =   3690
            Width           =   1395
            _extentx        =   2461
            _extenty        =   661
            caption         =   "Reset"
            font            =   "Tags.frx":02B4
         End
         Begin VB.Shape Shape4 
            Height          =   495
            Left            =   60
            Top             =   150
            Width           =   4905
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00E0E0E0&
            BorderWidth     =   2
            X1              =   0
            X2              =   1340
            Y1              =   0
            Y2              =   0
         End
      End
      Begin MMPlayerXProject.vkFrame frameBack 
         Height          =   5265
         Index           =   0
         Left            =   3540
         TabIndex        =   7
         Top             =   570
         Width           =   5340
         _extentx        =   9419
         _extenty        =   9287
         backcolor1      =   14737632
         font            =   "Tags.frx":02DC
         showtitle       =   0
         titlecolor1     =   12632256
         bordercolor     =   0
         borderwidth     =   0
         roundangle      =   0
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   4
            Left            =   90
            TabIndex        =   47
            Top             =   30
            Width           =   1395
            _extentx        =   2461
            _extenty        =   714
            backcolor1      =   16777215
            backcolor2      =   14345190
            backcolorpushed1=   14542053
            backcolorpushed2=   14345442
            caption         =   "ID3"
            font            =   "Tags.frx":0304
            forecolor       =   4210752
            bordercolor     =   7617536
            displaypicture  =   0
            drawfocus       =   0
            drawmouseinrect =   0
            disabledbackcolor=   15398133
            customstyle     =   0
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   5
            Left            =   1460
            TabIndex        =   46
            Top             =   30
            Width           =   1395
            _extentx        =   2461
            _extenty        =   714
            caption         =   "MORE"
            font            =   "Tags.frx":032C
            drawfocus       =   0
            drawmouseinrect =   0
            customstyle     =   0
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   6
            Left            =   2820
            TabIndex        =   45
            Top             =   30
            Width           =   1395
            _extentx        =   2461
            _extenty        =   714
            caption         =   "LYRICS"
            font            =   "Tags.frx":0354
            drawfocus       =   0
            drawmouseinrect =   0
            customstyle     =   0
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   1
            Left            =   330
            TabIndex        =   44
            Top             =   4720
            Width           =   1395
            _extentx        =   2461
            _extenty        =   714
            caption         =   "OK"
            font            =   "Tags.frx":037C
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   0
            Left            =   1980
            TabIndex        =   43
            Top             =   4720
            Width           =   1395
            _extentx        =   2461
            _extenty        =   714
            caption         =   "Cancel"
            font            =   "Tags.frx":03A4
         End
         Begin MMPlayerXProject.vkCommand vkBtn 
            Height          =   405
            Index           =   2
            Left            =   3540
            TabIndex        =   42
            Top             =   4720
            Width           =   1395
            _extentx        =   2461
            _extenty        =   714
            caption         =   "Apply"
            font            =   "Tags.frx":03CC
         End
         Begin MMPlayerXProject.vkFrame frameBack 
            Height          =   4245
            Index           =   3
            Left            =   90
            TabIndex        =   8
            Top             =   390
            Width           =   5160
            _extentx        =   9102
            _extenty        =   7594
            backcolor1      =   14474460
            font            =   "Tags.frx":03F4
            showtitle       =   0
            titlecolor1     =   12632256
            bordercolor     =   0
            roundangle      =   0
            Begin VB.TextBox txtTitle 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1560
               TabIndex        =   30
               Top             =   600
               Width           =   3270
            End
            Begin VB.TextBox txtArtist 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1560
               TabIndex        =   29
               Top             =   930
               Width           =   3270
            End
            Begin VB.TextBox txtAlbum 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               TabIndex        =   28
               Top             =   1255
               Width           =   3270
            End
            Begin VB.TextBox txtYear 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               MaxLength       =   4
               TabIndex        =   27
               Top             =   1555
               Width           =   540
            End
            Begin VB.TextBox txtTrack 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   26
               Top             =   270
               Width           =   585
            End
            Begin VB.TextBox txtComposer 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1560
               TabIndex        =   25
               Top             =   2865
               Width           =   3270
            End
            Begin VB.TextBox txtOrgArtist 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1560
               TabIndex        =   24
               Top             =   3195
               Width           =   3270
            End
            Begin VB.TextBox txtLink 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1560
               TabIndex        =   23
               Top             =   3525
               Width           =   3270
            End
            Begin VB.TextBox txtEncoder 
               Appearance      =   0  'Flat
               Height          =   315
               Left            =   1560
               TabIndex        =   22
               Top             =   3855
               Width           =   3270
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   1
               Left            =   1290
               TabIndex        =   21
               Top             =   1260
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":041C
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   2
               Left            =   1290
               TabIndex        =   20
               Top             =   1560
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":0444
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   3
               Left            =   2790
               TabIndex        =   19
               Top             =   1590
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":046C
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   4
               Left            =   1290
               TabIndex        =   18
               Top             =   1920
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":0494
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   5
               Left            =   1290
               TabIndex        =   17
               Top             =   2910
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":04BC
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   6
               Left            =   1290
               TabIndex        =   16
               Top             =   3210
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":04E4
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   7
               Left            =   1290
               TabIndex        =   15
               Top             =   3540
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":050C
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   8
               Left            =   1290
               TabIndex        =   14
               Top             =   3870
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":0534
            End
            Begin MMPlayerXProject.vkCheck chktags 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   0
               Left            =   1290
               TabIndex        =   13
               Top             =   960
               Width           =   225
               _extentx        =   397
               _extenty        =   450
               backcolor       =   16777215
               backstyle       =   0
               caption         =   ""
               font            =   "Tags.frx":055C
            End
            Begin VB.TextBox txtComment 
               Appearance      =   0  'Flat
               Height          =   990
               Left            =   1560
               MultiLine       =   -1  'True
               TabIndex        =   12
               Top             =   1860
               Width           =   3275
            End
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   3040
               ScaleHeight     =   255
               ScaleWidth      =   1755
               TabIndex        =   9
               Top             =   1560
               Width           =   1785
               Begin VB.ComboBox cboGenre 
                  Appearance      =   0  'Flat
                  Height          =   315
                  ItemData        =   "Tags.frx":0584
                  Left            =   -30
                  List            =   "Tags.frx":0586
                  Style           =   2  'Dropdown List
                  TabIndex        =   11
                  Top             =   -30
                  Width           =   1805
               End
               Begin VB.PictureBox Picture1 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00E0E0E0&
                  Enabled         =   0   'False
                  ForeColor       =   &H80000008&
                  Height          =   300
                  Left            =   1470
                  Picture         =   "Tags.frx":0588
                  ScaleHeight     =   270
                  ScaleWidth      =   300
                  TabIndex        =   10
                  Top             =   -10
                  Width           =   330
               End
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Year:"
               Height          =   195
               Index           =   1
               Left            =   795
               TabIndex        =   41
               Top             =   1575
               Width           =   465
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Artist:"
               Height          =   195
               Index           =   3
               Left            =   735
               TabIndex        =   40
               Top             =   975
               Width           =   525
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Album:"
               Height          =   195
               Index           =   4
               Left            =   645
               TabIndex        =   39
               Top             =   1275
               Width           =   615
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Title:"
               Height          =   195
               Index           =   5
               Left            =   825
               TabIndex        =   38
               Top             =   630
               Width           =   435
            End
            Begin VB.Label Label 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Genre:"
               Height          =   195
               Index           =   7
               Left            =   2160
               TabIndex        =   37
               Top             =   1605
               Width           =   600
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Comment:"
               Height          =   195
               Index           =   0
               Left            =   345
               TabIndex        =   36
               Top             =   1920
               Width           =   915
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Track:"
               Height          =   195
               Index           =   2
               Left            =   705
               TabIndex        =   35
               Top             =   300
               Width           =   555
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Composer:"
               Height          =   195
               Index           =   6
               Left            =   270
               TabIndex        =   34
               Top             =   2910
               Width           =   960
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Org. Artist:"
               Height          =   195
               Index           =   8
               Left            =   270
               TabIndex        =   33
               Top             =   3240
               Width           =   960
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Link:"
               Height          =   195
               Index           =   9
               Left            =   795
               TabIndex        =   32
               Top             =   3570
               Width           =   420
            End
            Begin VB.Label Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Encoder:"
               Height          =   195
               Index           =   10
               Left            =   450
               TabIndex        =   31
               Top             =   3900
               Width           =   765
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00E0E0E0&
               BorderWidth     =   2
               X1              =   15
               X2              =   1355
               Y1              =   0
               Y2              =   0
            End
         End
         Begin VB.Shape Shape1 
            Height          =   615
            Left            =   90
            Top             =   4610
            Width           =   5145
         End
      End
      Begin MMPlayerXProject.ucProgressBar pbProgress 
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   8685
         _extentx        =   15319
         _extenty        =   344
         font            =   "Tags.frx":0A02
         brushstyle      =   0
         color           =   12632256
         color2          =   16750899
         value           =   20
      End
      Begin VB.Label lblFile 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   90
         TabIndex        =   69
         Top             =   150
         Width           =   8715
      End
   End
End
Attribute VB_Name = "frmTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private FilesSelected As Integer
Private CurrentSelection As Integer
'// vars functions undo in lyrics
Private Arr() As Long
Private Const cChunk = 10
Private Last As Long, Cur As Long
Attribute Cur.VB_VarUserMemId = 1073938435
Dim Pos As Long
Attribute Pos.VB_VarUserMemId = 1073938437
Dim tabIndex As Integer
Attribute tabIndex.VB_VarUserMemId = 1073938438
Dim FilePlaying As String
Attribute FilePlaying.VB_VarUserMemId = 1073938439
Dim LastPosition As Long
Attribute LastPosition.VB_VarUserMemId = 1073938440
Dim LastState As String
Attribute LastState.VB_VarUserMemId = 1073938441
Dim iCurrentAlbum As Integer
Attribute iCurrentAlbum.VB_VarUserMemId = 1073938442
Dim bRestartPlayer As Boolean
Attribute bRestartPlayer.VB_VarUserMemId = 1073938443
Dim sLastGenre As String
Attribute sLastGenre.VB_VarUserMemId = 1073938444
Dim sArtFile As String
Attribute sArtFile.VB_VarUserMemId = 1073938445
Dim bLoadingTag As Boolean
Attribute bLoadingTag.VB_VarUserMemId = 1073938446
Dim InFormDrag As Boolean
Attribute InFormDrag.VB_VarUserMemId = 1073938447
Dim cWindows As New cWindowSkin
Attribute cWindows.VB_VarUserMemId = 1073938448
Dim cAjustarDesk As New clsDockingHandler
Attribute cAjustarDesk.VB_VarUserMemId = 1073938449
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Private Const SRCPAINT = &HEE0086
Private Const BLACKNESS = &H42
Private Const WHITENESS = &HFF0062
Private BackColor1 As Long
Attribute BackColor1.VB_VarUserMemId = 1073938450
Private BackColor2 As Long
Attribute BackColor2.VB_VarUserMemId = 1073938451
Private tabColor As Long
Attribute tabColor.VB_VarUserMemId = 1073938452
Public backColorSelect As Long
Attribute backColorSelect.VB_VarUserMemId = 1073938453
'////Following variables are for album art display
Private cImage As c32bppDIB
Attribute cImage.VB_VarUserMemId = 1073938454
Private tPic As IPictureDisp
Attribute tPic.VB_VarUserMemId = 1073938455
Private CurIndex As Long
Attribute CurIndex.VB_VarUserMemId = 1073938456
Private MaxIndex As Long
Attribute MaxIndex.VB_VarUserMemId = 1073938457
Private currentMp3file As String
Attribute currentMp3file.VB_VarUserMemId = 1073938458
Private bUpdatePicArt As Boolean
Attribute bUpdatePicArt.VB_VarUserMemId = 1073938459
Private Sub btnExit_Click()
' boolTagsLoaded = False
    boolTagsShow = False
    frmTags.Visible = False
    frmPopUp.mnuTagEditor.Checked = False
End Sub

Private Sub cboGenre_Click()
    If sLastGenre = "" Then Exit Sub
    If sLastGenre <> cboGenre.Text Then
        Update_Tags_Ref
        If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True
    End If
    sLastGenre = ""
End Sub

Private Sub cboGenre_DropDown()
    sLastGenre = cboGenre.Text
End Sub

Private Sub chkTags_Click(Index As Integer)
    Dim bolEnabled As Boolean
    Dim bChk As Boolean
    Dim i As Integer

    If chktags(Index).Value = vbChecked Then
        bolEnabled = True
    End If
    'Debug.Print FilesSelected
    Select Case Index
    Case 0    '// Artist
        txtArtist.Enabled = bolEnabled
    Case 1    '// Album
        txtAlbum.Enabled = bolEnabled
    Case 2    '// Year
        txtYear.Enabled = bolEnabled
    Case 3    '// genre
        cboGenre.Enabled = bolEnabled
    Case 4    '// Comment
        txtComment.Enabled = bolEnabled
    Case 5    '// Composer
        txtComposer.Enabled = bolEnabled
    Case 6    '//Org Artis
        txtOrgArtist.Enabled = bolEnabled
    Case 7    '// Link
        txtLink.Enabled = bolEnabled
    Case 8    '// Encoder
        txtEncoder.Enabled = bolEnabled
    Case 9    '// Art image
        vkBtn(7).Enabled = bolEnabled
        vkBtn(8).Enabled = bolEnabled
        ' If vkFiletags.CheckCount > 0 Then Label(11).Enabled = False
    End Select

    For i = 0 To 9
        If chktags(i).Value = vbChecked Then
            bChk = True
            Exit For
        End If
    Next
    vkBtn(2).Enabled = bChk  'APPLY
    'Debug.Print vkFiletags.CheckCount
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cmdAdd_Click
' Author    : Mahesh Kurmi
' Date      : 3/5/2011
' Purpose   : to add lyrics,initially it was a button but i have postponed working on it
' HOW       :
' Returns   :
' Parameters:
'---------------------------------------------------------------------------------------
'
Private Sub cmdAdd_Click()
'add a timestamp at the beginning of the current line (Lyrics)
    Dim OldMin As Long    'the minutes of old timestamp
    Dim OldSec As Long    'the seconds of old timestamp
    Dim oldHou As Long    'the hours of old timestamp
    Dim NewMin As Long    'the minutes of new timestamp
    Dim NewSec As Long    'the seconds of new timestamp
    Dim NewHou As Long    'the hours of new timestamp
    Dim LineLength As Long    'length of a line
    Dim CurrentLine As Long    'the current line number
    Dim TotalLines As Long    'how many lines there are
    Dim sCurrentTime As String    'the current time in string format
    Dim CharPos As Long    'character position

    Dim arryOldTime() As String
    Dim arryNewTime() As String
    Dim s As String, strTemp As String
    Dim j As Integer, Fin As Integer

    'error handler
    On Error GoTo hell
    '================================================================
    '  This is simple lyrics function
    '  how it work? good question :)
    '   - First load a file in tag editor
    '   - Write the lyrics in the text
    '   - Play the song with the over buttons
    '   - Use add button in the just time
    '                            is all, ¿Facil no?
    '================================================================

    If vkFiletags.ListCount = 0 Or PlayerState = "false" Then Exit Sub

    'check to make sure it contains a time
    sCurrentTime = Convert_Time_to_string(Stream_GetPosition(1))
    arryNewTime = Split(sCurrentTime, ":")

    'if has hours
    If UBound(arryNewTime) > 1 Then
        'convert the Time into integers
        NewHou = Val(arryNewTime(0))
        NewMin = Val(arryNewTime(1))
        NewSec = Val(arryNewTime(2))
    Else
        NewHou = 0
        NewMin = Val(arryNewTime(0))
        NewSec = Val(arryNewTime(1))
    End If

    'add the brackets to the time
    s = "[" & sCurrentTime & "]"

    'set the insert point to the beginning of the line, add 1 to it to make sure
    'we don't get a 0 length string compare.
    CurrentLine = SendMessage(txtLyrics.hwnd, EM_LINEFROMCHAR, txtLyrics.SelStart, ZERO)
    CharPos = SendMessage(txtLyrics.hwnd, EM_LINEINDEX, CurrentLine, ZERO)
    'get the length of the line
    LineLength = SendMessage(txtLyrics.hwnd, EM_LINELENGTH, CharPos, ZERO)
    LineLength = CharPos + LineLength
    Pos = CharPos + 1

    '// note: the [Do..Loop Until] is optional for look only
    '// you can delete and work lyrics function :P

    'check to make sure there is no timestamp already there, if so
    'then compare the new time to the old timestamp so the new one
    'is inserted at the correct point in end of old timestamp.

    'there is a timestamp here, get the time
    Do
        j = InStr(Pos, txtLyrics.Text, "[")
        If j > 0 And j <= LineLength Then
            Fin = InStr(Pos, txtLyrics.Text, "]")
            '// solo agregar letras hasta el formato 00:00:00
            If ((Fin - 1) - j) < 9 Then
                strTemp = Mid$(txtLyrics.Text, j + 1, Fin - j - 1)
            End If
        Else
            Exit Do
        End If

        arryOldTime = Split(strTemp, ":")

        'if has hours
        If UBound(arryOldTime) > 1 Then
            'convert the Time into integers
            oldHou = Val(arryOldTime(0))
            OldMin = Val(arryOldTime(1))
            OldSec = Val(arryOldTime(2))
        Else
            oldHou = 0
            OldMin = Val(arryOldTime(0))
            OldSec = Val(arryOldTime(1))
        End If

        'check to see if new timestamp is newer that old timestamp
        If (NewHou > oldHou) Or (NewHou = oldHou And NewMin > OldMin) Or (NewHou = oldHou And NewMin = OldMin And NewSec > OldSec) Then
            'yes, it is, so skip this one
            Pos = Fin + 1
        Else
            Exit Do
        End If
    Loop Until j = 0

    LineLength = 0

    'subtract one from the insert point and insert the stamp
    Pos = Pos - 1
    txtLyrics.SelStart = Pos
    txtLyrics.SelText = s
    'and push this position onto the undo stack
    Undo_Push Pos
    'enable the undo button
    cmdundo.Enabled = True

    'now drop them to the next non blank line, or back to the beginning
    'how many lines?
    TotalLines = SendMessage(txtLyrics.hwnd, EM_GETLINECOUNT, ZERO, ZERO)
    'safety check... should always be true
    If TotalLines > CurrentLine Then
        Do
            'increment current line
            CurrentLine = CurrentLine + 1
            'Get the position of the beginning of the line
            CharPos = SendMessage(txtLyrics.hwnd, EM_LINEINDEX, CurrentLine, ZERO)
            'get the length of the line
            LineLength = SendMessage(txtLyrics.hwnd, EM_LINELENGTH, CharPos, ZERO)
            'and keep looping until we get a non blank line or we get to the end
        Loop Until LineLength > 0 Or CurrentLine = TotalLines
        'if charpos = -1 then we are at the end.  Send them back to beginning
        If CharPos = -1 Then CharPos = 0
        'place cursor
        txtLyrics.SelStart = CharPos
    End If

    '/* update tags
    If Trim(txtLyrics.Text) <> "" Then Update_Tags_Ref

    'and set the focus back to the text box
    txtLyrics.SetFocus
    Exit Sub
hell:
End Sub

Private Sub Save_Tags()

    Dim strFileName As String
    Dim i As Integer
    Dim iCount As Integer
    Dim iFUpdated As Integer
    Dim cID3 As New cMP3
    Dim bChk As Boolean

    'On Error Resume Next
    '// if no checked all checkbox
    If vkFiletags.ListCount = 0 Then Exit Sub


    If FilesSelected > 1 Then
        For i = 0 To 9
            If chktags(i).Value = vbChecked Then
                bChk = True
                Exit For
            End If
        Next

        If bChk = False Then
            Exit Sub
        End If
    End If

    '// reset values for progress bar
    pbProgress.Min = 0
    pbProgress.Max = vkFiletags.ListCount
    pbProgress.Value = 0

    pbProgress.Visible = True
    cID3.Read_MPEGInfo = False
    For i = 0 To vkFiletags.ListCount - 1
        strFileName = listRef.ListItems(i + 1).Key

        DoEvents

        '// more than one files selected
        If FilesSelected > 1 Then

            If vkFiletags.Checked(i + 1) = True Then
                lblFile.Caption = "Updating file: " & listRef.ListItems(i + 1).Key
                cID3.Read_File_Tags strFileName
                '// make new tag

                '// Artist checked change at all
                If chktags(0).Value = vbChecked Then cID3.Artist = Trim(txtArtist.Text)

                '// Album checked change at all
                If chktags(1).Value = vbChecked Then cID3.Album = Trim(txtAlbum.Text)

                '// year checked change at all
                If chktags(2).Value = vbChecked Then cID3.Year = Trim(txtYear.Text)

                '// Genre checked change at all
                If chktags(3).Value = vbChecked Then cID3.Genre = cboGenre.Text

                '// Comments checked change at all
                If chktags(4).Value = vbChecked Then cID3.Comment = Trim(txtComment.Text)

                '// Composer checked change at all
                If chktags(5).Value = vbChecked Then cID3.Composer = Trim(txtComposer.Text)

                '// Orig. Artist checked change at all
                If chktags(6).Value = vbChecked Then cID3.OrigArtist = Trim(txtOrgArtist.Text)

                '// Links checked change at all
                If chktags(7).Value = vbChecked Then cID3.LinkTo = Trim(txtLink.Text)

                '// Encoder checked change at all
                If chktags(8).Value = vbChecked Then cID3.EncodedBy = Trim(txtEncoder.Text)

                '// Comments checked change at all
                If chktags(9).Value = vbChecked Then cID3.Images = sArtFile

                '// write the tags
                cID3.Write_File_Tags

                If LCase(strFileName) = LCase(sFileMainPlaying) Then
                    frmMain.Load_File_Tags
                    '  If bolLyricsShow = True Then frmMain.Start_Lyrics
                End If

                iFUpdated = iFUpdated + 1

            End If

        ElseIf Trim(listRef.ListItems(i + 1).SubItems(15)) = "simon" Then    '"simon" for if the track is to be edited or not
            cID3.Read_File_Tags strFileName
            '// make new tag
            cID3.TrackNr = Trim(listRef.ListItems.Item(i + 1).SubItems(1))
            cID3.Title = Trim(listRef.ListItems.Item(i + 1).SubItems(2))
            cID3.Artist = Trim(listRef.ListItems.Item(i + 1).SubItems(3))
            cID3.Album = Trim(listRef.ListItems.Item(i + 1).SubItems(4))
            cID3.Year = Trim(listRef.ListItems.Item(i + 1).SubItems(5))
            cID3.Genre = Trim(listRef.ListItems.Item(i + 1).SubItems(6))
            cID3.Comment = Trim(listRef.ListItems.Item(i + 1).SubItems(7))
            cID3.Composer = Trim(listRef.ListItems.Item(i + 1).SubItems(8))
            cID3.OrigArtist = Trim(listRef.ListItems.Item(i + 1).SubItems(9))
            cID3.LinkTo = Trim(listRef.ListItems.Item(i + 1).SubItems(10))
            cID3.EncodedBy = Trim(listRef.ListItems.Item(i + 1).SubItems(11))
            cID3.Images = Trim(listRef.ListItems.Item(i + 1).SubItems(12))
            cID3.Lyrics = Trim(listRef.ListItems.Item(i + 1).SubItems(13))

            '// write the tags
            cID3.Write_File_Tags

            '          sSQL = "UPDATE MUSIC SET TITLE='" & cID3.Title & "'," & _
                       '              "ARTIST='" & cID3.Artist & "',ALBUM='" & cID3.Album & "'," & _
                       '              "YEAR='" & cID3.Year & "',GENRE='" & cID3.Genre & "'" & _
                       '              " WHERE FILE='" & strFileName & "'"
            '            'Debug.Print sSQL
            '           CMD.CommandText = sSQL
            '           CMD.Execute

            If LCase(strFileName) = LCase(sFileMainPlaying) Then
                frmMain.Load_File_Tags
                ' If bolLyricsShow = True Then frmMain.Start_Lyrics
            End If

            listRef.ListItems.Item(i + 1).SubItems(1) = ""
            listRef.ListItems.Item(i + 1).SubItems(2) = ""
            listRef.ListItems.Item(i + 1).SubItems(3) = ""
            listRef.ListItems.Item(i + 1).SubItems(4) = ""
            listRef.ListItems.Item(i + 1).SubItems(5) = ""
            listRef.ListItems.Item(i + 1).SubItems(6) = ""
            listRef.ListItems.Item(i + 1).SubItems(7) = ""
            listRef.ListItems.Item(i + 1).SubItems(8) = ""
            listRef.ListItems.Item(i + 1).SubItems(9) = ""
            listRef.ListItems.Item(i + 1).SubItems(10) = ""
            listRef.ListItems.Item(i + 1).SubItems(11) = ""
            listRef.ListItems.Item(i + 1).SubItems(12) = ""
            listRef.ListItems.Item(i + 1).SubItems(13) = ""
            listRef.ListItems.Item(i + 1).SubItems(15) = ""

            'MODIFICAR LA LISTA DE REPRODUCCION CON LOS NUEVOS VALORES
            iFUpdated = iFUpdated + 1

        End If

        iCount = iCount + 1
        pbProgress.Value = iCount

    Next
    pbProgress.Visible = False
    lblFile.Caption = " Done! Updated [ " & iFUpdated & " ] files"
    Exit Sub
hell:
    MsgBox err.Description

End Sub







Private Sub vkcmdPlayer_Click(Index As Integer)

    If vkFiletags.ListCount = 0 Then Exit Sub

    Select Case Index
    Case 0    '// skip backward
        Five_Seg_Backward
    Case 1    '// play
        If frmMain.PlayerIsPlaying = "true" Then frmMain.Stop_Player
        If vkFiletags.ListIndex = -1 Then vkFiletags.ListIndex = 0

        Player_Play listRef.ListItems(vkFiletags.ListIndex + 1).Key
    Case 2    '// pause
        Pause_Play
    Case 3    '// stop
        Stop_Player
        FilePlaying = ""
    Case 4    '// skip forward
        Five_Seg_Forward
    End Select
    txtLyrics.SetFocus

End Sub



Private Sub cmdUndo_Click()
    Dim Fin As Integer, j As Integer, Start As Integer
    On Error GoTo hell

    With txtLyrics
        Start = Undo_Pop
        If Start = 0 Then Start = 1
        'select the timestamp
        j = InStr(Start, txtLyrics.Text, "[")
        If j > 0 Then
            Fin = InStr(Start + 1, txtLyrics.Text, "]")
            '// solo agregar letras hasta el formato 00:00:00
            If ((Fin - 1) - j) > 9 Then
                Fin = 0
            End If
        End If
        'get the postion of the last timestamp from the stack
        If Start = 1 Then Start = 0
        .SelStart = Start
        .SelLength = (Fin - Start)
        'and delete it
        .SelText = ""
        .SetFocus
    End With
    'If there is nothing in the stack, undo should not be enabled
    If Cur = 0 Then cmdundo.Enabled = False
    Exit Sub
hell:

End Sub

Private Sub Texts_Enableds(bolEnabled As Boolean)
    lblFile.Caption = ""
    chktags(0).Value = vbUnchecked
    chktags(1).Value = vbUnchecked
    chktags(2).Value = vbUnchecked
    chktags(3).Value = vbUnchecked
    chktags(4).Value = vbUnchecked
    chktags(5).Value = vbUnchecked
    chktags(6).Value = vbUnchecked
    chktags(7).Value = vbUnchecked
    chktags(8).Value = vbUnchecked
    chktags(9).Value = vbUnchecked

    chktags(0).Enabled = Not bolEnabled
    chktags(1).Enabled = Not bolEnabled
    chktags(2).Enabled = Not bolEnabled
    chktags(3).Enabled = Not bolEnabled
    chktags(4).Enabled = Not bolEnabled
    chktags(5).Enabled = Not bolEnabled
    chktags(6).Enabled = Not bolEnabled
    chktags(7).Enabled = Not bolEnabled
    chktags(8).Enabled = Not bolEnabled
    chktags(9).Enabled = Not bolEnabled

    chktags(0).Visible = Not bolEnabled
    chktags(1).Visible = Not bolEnabled
    chktags(2).Visible = Not bolEnabled
    chktags(3).Visible = Not bolEnabled
    chktags(4).Visible = Not bolEnabled
    chktags(5).Visible = Not bolEnabled
    chktags(6).Visible = Not bolEnabled
    chktags(7).Visible = Not bolEnabled
    chktags(8).Visible = Not bolEnabled
    chktags(9).Visible = Not bolEnabled
    Label(11).Visible = Not bolEnabled

    txtTrack.Enabled = bolEnabled
    txtTitle.Enabled = bolEnabled
    txtArtist.Enabled = bolEnabled
    txtAlbum.Enabled = bolEnabled
    txtYear.Enabled = bolEnabled
    cboGenre.Enabled = bolEnabled
    txtComment.Enabled = bolEnabled
    txtComposer.Enabled = bolEnabled
    txtOrgArtist.Enabled = bolEnabled
    txtLink.Enabled = bolEnabled
    txtEncoder.Enabled = bolEnabled
    Label(11).Enabled = bolEnabled
    txtLyrics.Text = ""
    lblMPEGInfo.Caption = ""


End Sub





Private Sub vkfiletags_Click()
    On Error Resume Next
    Dim i As Integer
    Dim tID3 As New cMP3

    If vkFiletags.ListCount = 0 Then Exit Sub
    FilesSelected = 0
    For i = 0 To vkFiletags.ListCount - 1
        If vkFiletags.Selected(i) = True Then
            FilesSelected = FilesSelected + 1
        End If
    Next i

    If PlayerState <> "false" Then Stop_Player

    '// pop for stack in lytics function
    Last = 10
    Cur = 0
    ReDim Arr(1 To Last) As Long
    cmdundo.Enabled = False

    If FilesSelected > 1 Then
        Texts_Enableds False
        lblFile.Caption = LineLanguage(61)
        vkBtn(2).Enabled = False
        Exit Sub
    Else
        Texts_Enableds True
    End If

    lblFile.Caption = listRef.ListItems(vkFiletags.ListIndex + 1).Key
    lblFile.ToolTipText = listRef.ListItems(vkFiletags.ListIndex + 1).Key
    tID3.Read_MPEGInfo = True
    tID3.Read_File_Tags listRef.ListItems(vkFiletags.ListIndex + 1).Key

    txtTrack.Text = tID3.TrackNr
    txtTitle.Text = tID3.Title
    txtAlbum.Text = tID3.Album
    txtArtist.Text = tID3.Artist
    txtYear.Text = tID3.Year

    For i = 0 To cboGenre.ListCount - 1
        If cboGenre.list(i) = tID3.Genre Then
            cboGenre.ListIndex = i
            Exit For
        End If
    Next i

    txtComment.Text = tID3.Comment
    txtComposer.Text = tID3.Composer
    txtOrgArtist.Text = tID3.OrigArtist
    txtLink.Text = tID3.LinkTo
    If tID3.LinkTo = "" Then txtLink.Text = "www.NPLayerx1.com"
    txtEncoder.Text = tID3.EncodedBy
    If tID3.EncodedBy = "" Then txtEncoder.Text = "MMPlayerX1 Version 2.0"
    txtLyrics.Text = tID3.Lyrics
    ' imgArt.Picture = LoadPicture(tID3.Images)
    ' sArtFile = tID3.Images


    lblMPEGInfo.Caption = "<> Size: " & tID3.MPEG_FileSizeMB & "  <> Length: " & tID3.MPEG_DurationTime & vbCrLf & _
                          "<> MPEG " & tID3.MPEG_Version & vbCrLf & _
                          "<> Bitrate: " & tID3.MPEG_Bit_Rate & " kbps, " & IIf(tID3.MPEG_VBR, "variable bit rate", "constant bit rate") & vbCrLf & _
                          "<> " & tID3.MPEG_Frequency & " Hz  " & tID3.MPEG_ChannelMode & vbCrLf & _
                          "<> CRCs: " & tID3.MPEG_CRCs & "  <> Copyrighted: " & tID3.MPEG_Copyrighted & vbCrLf & _
                          "<> Original: " & tID3.MPEG_Original & "  <> Emphasis: " & tID3.MPEG_Emphasis & vbCrLf & _
                          "<> ID3 v1 tag: " & tID3.TagID3V1 & vbCrLf & _
                          "<> ID3 v2 tag: " & tID3.TagID3V2
    Set tID3 = Nothing
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim i As Integer
    Dim ID3 As New cMP3

    tabIndex = 4
    boolTagsShow = True
    boolTagsLoaded = True
    frameBackmain.Top = 37
    frameBackmain.Left = 21
    frameBackmain.Height = 396
    frameBackmain.Width = 598
    cboGenre.Appearance = 0
    'Me.Width = 9510
    'Me.Height = 6975
    Set cImage = New c32bppDIB
    cImage.InitializeDIB ScaleX(PicAlbumArt.ScaleWidth, PicAlbumArt.ScaleMode, vbPixels), ScaleY(PicAlbumArt.ScaleHeight, PicAlbumArt.ScaleMode, vbPixels)

    '// initialize values for undo functions
    Last = 10
    Cur = 0
    ReDim Arr(1 To Last) As Long

    For i = 0 To 147
        cboGenre.AddItem ID3.GetGenreName(i)
    Next i
    boolTagsShow = True


    ' Me.Icon = frmMain.Icon
    frmTags.Left = (Screen.Width - frmTags.Width) / 2
    frmTags.Top = (Screen.Height - frmTags.Height) / 2
    'cImage.LoadPicture_File "c:\no_image.gif"
    LoadSkin

    'cImage.Render PicAlbumArt.hDC, 0, 0, PicAlbumArt.ScaleWidth, PicAlbumArt.ScaleHeight
    PlayerState = "false"
    Set ID3 = Nothing
    cImage.CreateCheckerBoard 8, vbWhite, PicAlbumArt.BackColor
    'cImage.DrawText_stdFont Me.Font, "TEXT OVERLAY" & vbNewLine & "Example #3", , 39, cImage.Height \ 2 + 1, , , vbBlack, , , , -90, True

    ' Load_Language_Tags
    PicAlbumArt.AutoRedraw = True
    cImage.Render PicAlbumArt.hDC
    PicAlbumArt.Picture = PicAlbumArt.Image

End Sub

Sub Load_Tags(sPath As String, Optional Row As Integer = -1)
    On Error GoTo hell

    '????????WHY-ERROR???????
    'Error occurs when the track with same path is already added
    'We dont want track with same path appear twice in list else it will create ambiguity in tag editing with multiple files
    '????????????????????

    Dim sFilename As String
    Dim aFile() As String

    aFile = Split(sPath, "\", , vbTextCompare)
    sFilename = aFile(UBound(aFile))
    If Not FileExists(sPath) Then Exit Sub

    listRef.ListItems.Add , sPath, sFilename
    listRef.ListItems.Item(sPath).SubItems(14) = sPath
    vkFiletags.AddItem sFilename

    'IF YOU SEND PLAY LIST Then row <> -1
    If Row <> -1 Then
        listRef.ListItems.Item(sPath).SubItems(16) = Row
    End If
hell:
End Sub
Private Sub Player_Play(FilePlay As String)

    On Error GoTo error
    If PlayerState = "pause" Then Pause_Play: Exit Sub

    Stream_Open FilePlay, FSOUND_NORMAL, 1, True, frmMain.VolumeNActuaL

    '// volume in main form
    'Stream_SetVolume 1, frmMain.VolumeNActuaL
    PlayerState = "true"
    FilePlaying = FilePlay
    Exit Sub
error:
    PlayerState = "false"
    FilePlaying = ""
    Stop_Player
End Sub

Sub Stop_Player()
    On Error Resume Next

    If vkFiletags.ListCount = 0 Then Exit Sub

    Stream_Stop 1
    PlayerState = "false"
End Sub

Private Sub Pause_Play()
    Dim CurState As Long

    On Error Resume Next

    If vkFiletags.ListCount = 0 Then Exit Sub

    If PlayerState = "false" Then Exit Sub
    CurState = Stream_GetState(1)
    '------'Esta Reproduciendo, pausar-------------------------------------------
    If CurState = 2 Then
        Stream_Pause 1
        PlayerState = "pause"
    Else
        '------'Si esta pausado, reproducir---------------------------------------------
        Stream_Pause 1
        PlayerState = "true"
    End If
End Sub

Sub Five_Seg_Forward()
    On Error GoTo hell
    Dim CurPos As Long

    If vkFiletags.ListCount = 0 Or PlayerState = "false" Then Exit Sub
    If PlayerState = "pause" Then Pause_Play

    CurPos = Stream_GetPosition(1)
    CurPos = CurPos + 5
    If CurPos > Stream_GetDuration(1) Then CurPos = Stream_GetDuration(1)
    Stream_SetPosition 1, CurPos
    Exit Sub
hell:

End Sub

'+-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-+
Sub Five_Seg_Backward()
    On Error GoTo hell
    Dim CurPos As Long
    If vkFiletags.ListCount = 0 Or PlayerState = "false" Then Exit Sub
    If PlayerState = "pause" Then Pause_Play
    CurPos = Stream_GetPosition(1)
    CurPos = CurPos - 5
    If CurPos < 0 Then CurPos = 0
    Stream_SetPosition 1, CurPos
    Exit Sub
hell:
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If PlayerState <> "false" Then Stop_Player
    boolTagsShow = False
    boolTagsLoaded = False
    cImage.DestroyDIB

End Sub




'//------------------------------------------------------------------------------//
'// functions for undo function in lyrics
Private Sub Undo_Push(Arg As Long)
    Cur = Cur + 1
    On Error GoTo FailPush
    Arr(Cur) = Arg
    Exit Sub
FailPush:
    Last = Last + cChunk  ' Grow
    ReDim Preserve Arr(1 To Last) As Long
    Resume                  ' Try again
End Sub

Private Function Undo_Pop() As Long
    If Cur Then
        Undo_Pop = Arr(Cur)
        Cur = Cur - 1
        If Cur < (Last - cChunk) Then
            Last = Last - cChunk      ' Shrink
            ReDim Preserve Arr(1 To Last) As Long
        End If
    End If
End Function


Private Sub Update_Tags_Ref()
    If vkFiletags.ListCount = 0 Then Exit Sub
    If FilesSelected > 1 Then Exit Sub
    With listRef.ListItems.Item(CurrentSelection)
        .SubItems(1) = txtTrack.Text
        .SubItems(2) = txtTitle.Text
        .SubItems(3) = txtArtist.Text
        .SubItems(4) = txtAlbum.Text
        .SubItems(5) = txtYear.Text
        .SubItems(6) = cboGenre.list(cboGenre.ListIndex)
        .SubItems(7) = txtComment.Text
        .SubItems(8) = txtComposer.Text
        .SubItems(9) = txtOrgArtist.Text
        .SubItems(10) = txtLink.Text
        .SubItems(11) = txtEncoder.Text
        .SubItems(12) = sArtFile
        .SubItems(13) = txtLyrics.Text
        .SubItems(15) = "simon"
        .SubItems(16) = str(CurrentSelection)
    End With
End Sub



Private Sub txtAlbum_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub

Private Sub txtArtist_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True
End Sub


Private Sub txtComment_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    'Debug.Print KeyCode
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub

Private Sub txtLyrics_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub

Private Sub txtTitle_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True
End Sub

Private Sub txtYear_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub
Private Sub txtTrack_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub

Private Sub txtComposer_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub

Private Sub txtOrgArtist_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub

Private Sub txtLink_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub

Private Sub txtEncoder_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Or KeyCode = 16 Or KeyCode = 17 Or KeyCode = 20 Or _
       KeyCode = 37 Or KeyCode = 38 Or KeyCode = 40 Or KeyCode = 39 Then Exit Sub
    Update_Tags_Ref
    If vkBtn(2).Enabled = False Then vkBtn(2).Enabled = True

End Sub


Public Sub LoadSkin()
    On Error Resume Next
    Dim iX, iY As Integer
    Set cWindows.FormularioPadre = Me
    Set cAjustarDesk.ParentForm = Me
    Me.Height = Read_INI("FORM", "formheight", 6975, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    Me.Width = Read_INI("FORM", "formwidth", 9610, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    cWindows.ColorInvisible = Read_INI("NormalMode", "ColorTrans", RGB(255, 0, 255), True)
    iX = Read_INI("Configuration", "ExitButtonX", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    iY = Read_INI("Configuration", "ExitButtonY", 1, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    cWindows.ButtonExitXY CLng(iX), CLng(iY)
    cWindows.MinimoAlto = Read_INI("Configuration", "MinHeight", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    cWindows.MinimoAncho = Read_INI("Configuration", "MinWidth", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    'iButtonsLeft = Read_INI("Configuration", "ButtonsLeft", 5, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    'iButtonsTop = Read_INI("Configuration", "ButtonsTop", 10, False, False, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\PLAY_LIST\config.ini")
    cWindows.CargarSkin tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\", True
    Picture1.Picture = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\combo.bmp")
    Dim k
    k = Read_Config_Button(btnExit, "Configuration", "exitButton", "0,0,10,10", tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    Set btnExit.PictureNormal = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\Exitnormal.bmp")
    Set btnExit.PictureOver = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\Exitover.bmp")
    Set btnExit.PictureDown = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\ExitDown.bmp")
    'Set btnExit.PictureNormal = LoadPicture(tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\DSP\Exitnormal.bmp")
    LoadConfig
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, 161, 2, 0


End Sub

Private Sub vkBtn_Click(Index As Integer)
    Dim sFile As String, sfilter As String
    Dim newWidth As Long, newHeight As Long

    Select Case Index
    Case 0:    'cancel
        Unload Me

    Case 1:    'ok
        vkBtn(1).Enabled = False
        If vkBtn(2).Enabled = True Then Save_Tags
        Unload Me

    Case 2:    'apply
        FilesSelected = FilesSelected
        Save_Tags
        vkBtn(2).Enabled = False

    Case 3:    'Select all
        If vkFiletags.ListCount <= 0 Then Exit Sub
        If vkBtn(Index).Tag <> "None" Then
            vkBtn(Index).Tag = "None"
            vkBtn(Index).Caption = "Select None"
            vkFiletags.CheckAll
            FilesSelected = vkFiletags.ListCount
        Else
            vkBtn(Index).Tag = "All"
            vkBtn(Index).Caption = "Select All"
            vkFiletags.UnCheckAll
            FilesSelected = 0
        End If
        If vkFiletags.CheckCount > 1 Then
            Texts_Enableds False
            lblFile.Caption = LineLanguage(61)
            vkBtn(2).Enabled = False
            Exit Sub
        Else
            Texts_Enableds True
        End If

    Case 4:    'Tab1
        frameBack(3).ZOrder vbBringToFront
        frameBack(4).ZOrder vbSendToBack
        frameBack(5).ZOrder vbSendToBack
        tabIndex = 4

    Case 5:    'tab2
        frameBack(4).ZOrder vbBringToFront
        tabIndex = 5
        vkBtn(7).Enabled = IIf((currentMp3file = "") Or (vkFiletags.CheckCount > 0 And chktags(9).Value = vbUnchecked), False, True)

    Case 6:    'Tab3
        frameBack(5).ZOrder vbBringToFront
        tabIndex = 6

    Case 7:    'Add art
        sfilter = "Image Files (*jpg, *.bmp, *.gif,*.png)|*.jpg;*.bmp;*.gif;*.png;*.ico"
        sFile = clsDlg.GetOpenAsName(frmTags.hwnd, "Open Album Art", , sfilter)
        If Dir(sFile) = "" Or sFile = "" Then Exit Sub

        PicAlbumArt.cls
        sArtFile = sFile
        cImage.LoadPicture_File sArtFile
        cImage.ScaleImage PicAlbumArt.ScaleWidth, PicAlbumArt.ScaleHeight, newWidth, newHeight, scaleDownAsNeeded
        cImage.Render PicAlbumArt.hDC, , , PicAlbumArt.ScaleWidth, PicAlbumArt.ScaleHeight
        ' Skip upadation if more than one files are selected in list box
        If vkFiletags.CheckCount > 0 Then vkUpDownArt.Visible = False: Exit Sub
        bUpdatePicArt = False    'make sure that picture box isnot repainted
        Set tPic = PicAlbumArt.Image
        If cGetInputState() <> 0 Then DoEvents

        If WriteAlbumArt(currentMp3file, CurIndex, tPic) Then
            CurIndex = CurIndex + 1
            MaxIndex = MaxIndex + 1
            vkBtn(11).Enabled = True
            vkBtn(8).Enabled = True
            vkUpDownArt.Max = MaxIndex
            vkUpDownArt.Value = CurIndex
            If MaxIndex > 1 Then vkUpDownArt.Visible = True: bUpdatePicArt = True
        End If

        '////////used Later 18-1-11
        'Update_Tags_Ref
        '/////////

    Case 8:    'remove art
        Dim k As Long
        bUpdatePicArt = False    'make sure that picture box isnot repainted
        If Dir(currentMp3file) = "" Or MaxIndex <= 0 Then Exit Sub
        ' Skip upadation if more than one files are selected in list box
        If vkFiletags.CheckCount > 0 Then PicAlbumArt.cls: vkUpDownArt.Visible = False: Exit Sub
        If DeleteAlbumArt(currentMp3file, CurIndex) Then
            PicAlbumArt.cls
            MaxIndex = MaxIndex - 1
            vkUpDownArt.Max = MaxIndex

            If MaxIndex <= 1 Then vkUpDownArt.Visible = False
            If CurIndex - 1 > 0 Then
                CurIndex = CurIndex - 1
                If ReadAlbumArt(currentMp3file, CurIndex, tPic, k) Then
                    ResizePic
                End If
            ElseIf MaxIndex > 0 Then
                If ReadAlbumArt(currentMp3file, CurIndex, tPic, k) Then
                    ResizePic
                End If
            Else
                CurIndex = 0
                vkBtn(11).Enabled = False
                vkBtn(8).Enabled = False
                Exit Sub
            End If
        End If

        vkUpDownArt.Value = CurIndex
        bUpdatePicArt = True

    Case 9:    'Add Lyrics
        Add_lyrics

    Case 11:    'Save art
        sfilter = "Image Files (*jpg, *.bmp, *.gif,*.png)|*.jpg;*.bmp;*.gif;*.png;*.ico"
        sFile = ""
        sFile = clsDlg.GetSaveAsName(frmTags.hwnd, "Save Album Art", , sfilter, "*.jpg")
        If sFile <> "" Then Call cImage.SaveToFile(sFile, True)

    End Select


    Dim i As Integer
    If Index <= 3 Or Index > 6 Then Exit Sub

    For i = 4 To 6
        vkBtn(i).BackColor2 = BackColor1
        vkBtn(i).BackColor1 = BackColor1    'vbWhite 'rgb(210,230,250)backColor1 'vbWhite
    Next
    vkBtn(Index).BackColor2 = BackColor1
    vkBtn(Index).BackColor1 = tabColor    'rgb(210,230,250)backColor1 'vbWhite

End Sub


Private Sub LoadConfig()
    Dim i As Integer
    Dim ListScroll As New vkPrivateScroll

    BackColor1 = cRead_INI("BUTTON", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    BackColor2 = cRead_INI("BUTTON", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    backColorSelect = cRead_INI("BUTTON", "BackcolorSelect", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    tabColor = cRead_INI("BUTTON", "tabColor", vbWhite, tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    frameBackmain.Top = Read_INI("CONTAINER", "top", 12, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    frameBackmain.Left = Read_INI("CONTAINER", "left", 12, , , tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    frameBackmain.BackColor1 = cRead_INI("CONTAINER", "Backcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    frameBackmain.BackColor2 = cRead_INI("CONTAINER", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")

    For i = 0 To 11
        vkBtn(i).BackColor1 = BackColor1
        vkBtn(i).BackColor2 = BackColor2
        vkBtn(i).ForeColor = cRead_INI("BUTTON", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
        vkBtn(i).BorderColor = cRead_INI("BUTTON", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
        vkBtn(i).BackColorPushed1 = cRead_INI("BUTTON", "BackColorPushed1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
        vkBtn(i).BackColorPushed2 = cRead_INI("BUTTON", "BackColorPushed2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    Next

    For i = 0 To 4
        vkcmdPlayer(i).BackColor1 = cRead_INI("BUTTON", "cbackcolor1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")    'BackColor1
        vkcmdPlayer(i).BackColor2 = cRead_INI("BUTTON", "cbackcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")    'BackColor2
        vkcmdPlayer(i).ForeColor = cRead_INI("BUTTON", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
        vkcmdPlayer(i).BorderColor = cRead_INI("BUTTON", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
        vkcmdPlayer(i).BackColorPushed1 = cRead_INI("BUTTON", "cBackColorPushed1", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
        vkcmdPlayer(i).BackColorPushed2 = cRead_INI("BUTTON", "cBackColorPushed2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")

    Next

    pbProgress.Color = BackColor1
    Line1.BorderColor = BackColor1
    Line2.BorderColor = BackColor1
    Line3.BorderColor = BackColor1

    Line1.X1 = vkBtn(4).Left - 80
    Line1.X2 = vkBtn(4).Left + vkBtn(4).Width - 140
    Line2.X1 = vkBtn(5).Left - 80
    Line2.X2 = vkBtn(5).Left + vkBtn(5).Width - 140
    Line3.X1 = vkBtn(6).Left - 80
    Line3.X2 = vkBtn(6).Left + vkBtn(6).Width - 140

    For i = 0 To 5
        frameBack(i).BackColor2 = cRead_INI("CONTAINER", "Backcolor2", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
        frameBack(i).BackColor1 = cRead_INI("CONTAINER", "Backcolor1", RGB(80, 110, 160), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    Next

    Dim txtBackcolor, txtForecolor
    txtBackcolor = cRead_INI("TEXT", "Backcolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    txtTrack.BackColor = txtBackcolor
    txtTitle.BackColor = txtBackcolor
    txtArtist.BackColor = txtBackcolor
    txtAlbum.BackColor = txtBackcolor
    txtYear.BackColor = txtBackcolor
    txtComment.BackColor = txtBackcolor
    txtComposer.BackColor = txtBackcolor
    txtOrgArtist.BackColor = txtBackcolor
    txtLink.BackColor = txtBackcolor
    txtEncoder.BackColor = txtBackcolor
    cboGenre.BackColor = txtBackcolor
    txtLyrics.BackColor = txtBackcolor

    txtForecolor = cRead_INI("TEXT", "forecolor", RGB(0, 60, 115), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    txtTrack.ForeColor = txtForecolor
    txtTitle.ForeColor = txtForecolor
    txtArtist.ForeColor = txtForecolor
    txtAlbum.ForeColor = txtForecolor
    txtYear.ForeColor = txtForecolor
    txtComment.ForeColor = txtForecolor
    txtComposer.ForeColor = txtForecolor
    txtOrgArtist.ForeColor = txtForecolor
    txtLink.ForeColor = txtForecolor
    txtEncoder.ForeColor = txtForecolor
    cboGenre.ForeColor = txtForecolor
    txtLyrics.ForeColor = txtForecolor

    PicAlbumArt.BackColor = frameBack(1).BackColor1
    cImage.Render PicAlbumArt.hDC, , , PicAlbumArt.ScaleWidth, PicAlbumArt.ScaleHeight
    For i = 0 To 14
        Label(i).ForeColor = txtForecolor
    Next

    lblMPEGInfo.ForeColor = txtForecolor
    lblFile.ForeColor = txtForecolor

    ListScroll.BackColor = cRead_INI("LIST", "ScrollbackColor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    ListScroll.FrontColor = cRead_INI("LIST", "ScrollfrontColor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    ListScroll.BorderColor = cRead_INI("LIST", "ScrollborderColor", RGB(0, 0, 0), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    ListScroll.ArrowColor = ListScroll.BorderColor
    vkFiletags.ForeColor = cRead_INI("LIST", "forecolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    vkFiletags.BorderColor = cRead_INI("LIST", "bordercolor", RGB(70, 100, 150), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    vkFiletags.BackColor = cRead_INI("LIST", "backcolor", RGB(70, 100, 50), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    vkFiletags.SelColor = cRead_INI("LIST", "selcolor", RGB(70, 100, 50), tAppConfig.AppConfig & "Skins\" & tAppConfig.Skin & "\IDE\config.ini")
    vkFiletags.VScroll = ListScroll

    Call vkBtn_Click(tabIndex)
End Sub


Private Sub vkFiletags_ItemChek(Item As vkListItem)
    If Item Is Nothing Then Exit Sub
    FilesSelected = vkFiletags.CheckCount

    If FilesSelected = vkFiletags.ListCount Then
        vkBtn(3).Tag = "None"
        vkBtn(3).Caption = "Select None"
    End If

    'Debug.Print FilesSelected

    If PlayerState <> "false" Then Stop_Player
    '// pop for stack in lytics function
    Last = 10
    Cur = 0
    ReDim Arr(1 To Last) As Long
    cmdundo.Enabled = False

    If FilesSelected > 0 Then
        Texts_Enableds False
        lblFile.Caption = LineLanguage(61)
        vkBtn(2).Enabled = False
        Exit Sub
    Else
        Texts_Enableds True
    End If

End Sub

Private Sub vkFiletags_ItemClick(Item As vkListItem)
'On Error Resume Next
    If Item Is Nothing Then Exit Sub
    If vkFiletags.ListCount = 0 Then Exit Sub
    CurrentSelection = Item.Index
    Call Show_tags(Item.Index)

End Sub

Private Sub vkFiletags_ItemUnCheck(Item As vkListItem)

    If FilesSelected = vkFiletags.ListCount Then
        vkBtn(3).Tag = "All"
        vkBtn(3).Caption = "Select All"
    End If

    FilesSelected = vkFiletags.CheckCount
    'Debug.Print FilesSelected

    If PlayerState <> "false" Then Stop_Player
    '// pop for stack in lytics function
    Last = 10
    Cur = 0
    ReDim Arr(1 To Last) As Long
    cmdundo.Enabled = False

    If FilesSelected > 0 Then
        Texts_Enableds False
        lblFile.Caption = LineLanguage(61)
        vkBtn(2).Enabled = False
        Exit Sub
    Else
        Texts_Enableds True
    End If
End Sub

Public Sub Show_tags(Index As Integer)
    Dim tID3 As New cMP3
    Dim i As Integer
    Dim k As Long
    lblFile.Caption = listRef.ListItems(Index).Key
    currentMp3file = IIf(FileExists(lblFile.Caption), lblFile.Caption, "")
    lblFile.ToolTipText = currentMp3file
    If currentMp3file = "" Then vkBtn(7).Enabled = False: Exit Sub
    vkBtn(7).Enabled = IIf(vkFiletags.CheckCount = 0, True, False)

    tID3.Read_MPEGInfo = True
    tID3.Read_File_Tags currentMp3file

    txtTrack.Text = tID3.TrackNr
    txtTitle.Text = tID3.Title
    txtAlbum.Text = tID3.Album
    txtArtist.Text = tID3.Artist
    txtYear.Text = tID3.Year

    For i = 0 To cboGenre.ListCount - 1
        If cboGenre.list(i) = tID3.Genre Then
            cboGenre.ListIndex = i
            Exit For
        End If
    Next i

    txtComment.Text = tID3.Comment
    txtComposer.Text = tID3.Composer
    txtOrgArtist.Text = tID3.OrigArtist
    txtLink.Text = tID3.LinkTo
    If tID3.LinkTo = "" Then txtLink.Text = "www.MMPlayerx1.com"
    txtEncoder.Text = tID3.EncodedBy
    If tID3.EncodedBy = "" Then txtEncoder.Text = "MMPlayerX1 Version 2.0"
    txtLyrics.Text = tID3.Lyrics


    lblMPEGInfo.Caption = "<> Size: " & tID3.MPEG_FileSizeMB & "  <> Length: " & tID3.MPEG_DurationTime & vbCrLf & _
                          "<> MPEG " & tID3.MPEG_Version & vbCrLf & _
                          "<> Bitrate: " & tID3.MPEG_Bit_Rate & " kbps, " & IIf(tID3.MPEG_VBR, "variable bit rate", "constant bit rate") & vbCrLf & _
                          "<> " & tID3.MPEG_Frequency & " Hz  " & tID3.MPEG_ChannelMode & vbCrLf & _
                          "<> CRCs: " & tID3.MPEG_CRCs & "  <> Copyrighted: " & tID3.MPEG_Copyrighted & vbCrLf & _
                          "<> Original: " & tID3.MPEG_Original & "  <> Emphasis: " & tID3.MPEG_Emphasis & vbCrLf & _
                          "<> ID3 v1 tag: " & tID3.TagID3V1 & vbCrLf & _
                          "<> ID3 v2 tag: " & tID3.TagID3V2
    Set tID3 = Nothing

    'imgArt.Picture = LoadPicture(tID3.Images) 'Not working
    'sArtFile = tID3.Images                    'Not working
    '////CODE for album Art

    PicAlbumArt.cls
    vkUpDownArt.Visible = False
    bUpdatePicArt = False
    If ID3Exist(currentMp3file) Then
        MaxIndex = GetAlbumArtCount(currentMp3file)
        If MaxIndex > 0 Then
            CurIndex = 1
            vkBtn(11).Enabled = IIf(vkFiletags.CheckCount = 0, True, False)
            vkBtn(8).Enabled = vkBtn(11).Enabled
            If ReadAlbumArt(currentMp3file, CurIndex, tPic, k) Then
                ResizePic
            End If
        Else
            CurIndex = 0
            vkBtn(11).Enabled = False
            vkBtn(8).Enabled = False
        End If
    Else
        CurIndex = 0
        MaxIndex = 0
        vkBtn(11).Enabled = False
        vkBtn(8).Enabled = False
    End If
    vkUpDownArt.Min = 1
    vkUpDownArt.Value = 1
    vkUpDownArt.Max = MaxIndex
    'Debug.Print MaxIndex & "sss" & vkUpDownArt.Max

    If MaxIndex > 1 Then vkUpDownArt.Visible = True: bUpdatePicArt = True

    '////Code for Album art
End Sub

Private Sub ResizePic()
'load Tpic image to cImage
    PicAlbumArt.cls
    Call cImage.LoadPicture_StdPicture(tPic)
    'Load Cimage to Picturebox
    cImage.ScaleImage PicAlbumArt.ScaleWidth, PicAlbumArt.ScaleHeight, PicAlbumArt.ScaleWidth, PicAlbumArt.ScaleHeight, scaleDownAsNeeded
    cImage.Render PicAlbumArt.hDC, , , PicAlbumArt.ScaleWidth, PicAlbumArt.ScaleHeight

    '////////////////////////////////////////////
    'The following code loades albumart directly in pcturebox (without Cimage)        'Rendering directly (but distorts image)
    'Dim nWidth  As Long
    'Dim nHeight As Long

    'On Error Resume Next
    'nWidth = ScaleX(TPic.Width, vbHimetric, vbPixels)
    'nHeight = ScaleY(TPic.Height, vbHimetric, vbPixels)
    'With PicAlbumArt
    '   If .ScaleWidth < (nWidth * (.ScaleHeight / nHeight)) Then
    '      nHeight = nHeight * (.ScaleWidth / nWidth)
    '     nWidth = .ScaleWidth
    ' Else
    '    nWidth = nWidth * (.ScaleHeight / nHeight)
    '   nHeight = .ScaleHeight
    'End If
    ' TPic.Render .hDC, (.ScaleWidth - CLng(nWidth)) / 2, (.ScaleHeight - CLng(nHeight)) / 2, CLng(nWidth), CLng(nHeight), 0, TPic.Height, TPic.Width, -TPic.Height, ByVal 0&
    'End With
    '//////////////////////////////
End Sub

Public Sub Add_lyrics()
'add a timestamp at the beginning of the current line (Lyrics)

    Dim OldMin As Long    'the minutes of old timestamp
    Dim OldSec As Long    'the seconds of old timestamp
    Dim oldHou As Long    'the hours of old timestamp
    Dim NewMin As Long    'the minutes of new timestamp
    Dim NewSec As Long    'the seconds of new timestamp
    Dim NewHou As Long    'the hours of new timestamp
    Dim LineLength As Long    'length of a line
    Dim CurrentLine As Long    'the current line number
    Dim TotalLines As Long    'how many lines there are
    Dim sCurrentTime As String    'the current time in string format
    Dim CharPos As Long    'character position

    Dim arryOldTime() As String
    Dim arryNewTime() As String
    Dim s As String, strTemp As String
    Dim j As Integer, Fin As Integer

    'error handler
    On Error GoTo hell
    '================================================================
    '  This is simple lyrics function
    '  how it work? good question :)
    '   - First load a file in tag editor
    '   - Write the lyrics in the text
    '   - Play the song with the over buttons
    '   - Use add button in the just time
    '                            is all, ¿Facil no?
    '================================================================

    If vkFiletags.ListCount = 0 Or PlayerState = "false" Then Exit Sub

    'check to make sure it contains a time
    sCurrentTime = Convert_Time_to_string(Stream_GetPosition(1))
    arryNewTime = Split(sCurrentTime, ":")

    'if has hours
    If UBound(arryNewTime) > 1 Then
        'convert the Time into integers
        NewHou = Val(arryNewTime(0))
        NewMin = Val(arryNewTime(1))
        NewSec = Val(arryNewTime(2))
    Else
        NewHou = 0
        NewMin = Val(arryNewTime(0))
        NewSec = Val(arryNewTime(1))
    End If

    'add the brackets to the time
    s = "[" & sCurrentTime & "]"

    'set the insert point to the beginning of the line, add 1 to it to make sure
    'we don't get a 0 length string compare.
    CurrentLine = SendMessage(txtLyrics.hwnd, EM_LINEFROMCHAR, txtLyrics.SelStart, ZERO)
    CharPos = SendMessage(txtLyrics.hwnd, EM_LINEINDEX, CurrentLine, ZERO)
    'get the length of the line
    LineLength = SendMessage(txtLyrics.hwnd, EM_LINELENGTH, CharPos, ZERO)
    LineLength = CharPos + LineLength
    Pos = CharPos + 1

    '// note: the [Do..Loop Until] is optional for look only
    '// you can delete and work lyrics function :P

    'check to make sure there is no timestamp already there, if so
    'then compare the new time to the old timestamp so the new one
    'is inserted at the correct point in end of old timestamp.

    'there is a timestamp here, get the time
    Do
        j = InStr(Pos, txtLyrics.Text, "[")
        If j > 0 And j <= LineLength Then
            Fin = InStr(Pos, txtLyrics.Text, "]")
            '// just add letters to the format 00:00:00
            If ((Fin - 1) - j) < 9 Then
                strTemp = Mid$(txtLyrics.Text, j + 1, Fin - j - 1)
            End If
        Else
            Exit Do
        End If

        arryOldTime = Split(strTemp, ":")

        'if has hours
        If UBound(arryOldTime) > 1 Then
            'convert the Time into integers
            oldHou = Val(arryOldTime(0))
            OldMin = Val(arryOldTime(1))
            OldSec = Val(arryOldTime(2))
        Else
            oldHou = 0
            OldMin = Val(arryOldTime(0))
            OldSec = Val(arryOldTime(1))
        End If

        'check to see if new timestamp is newer that old timestamp
        If (NewHou > oldHou) Or (NewHou = oldHou And NewMin > OldMin) Or (NewHou = oldHou And NewMin = OldMin And NewSec > OldSec) Then
            'yes, it is, so skip this one
            Pos = Fin + 1
        Else
            Exit Do
        End If
    Loop Until j = 0
    LineLength = 0

    'subtract one from the insert point and insert the stamp
    Pos = Pos - 1
    txtLyrics.SelStart = Pos
    txtLyrics.SelText = s
    'and push this position onto the undo stack
    Undo_Push Pos
    'enable the undo button
    vkBtn(8).Enabled = True

    'now drop them to the next non blank line, or back to the beginning
    'how many lines?
    TotalLines = SendMessage(txtLyrics.hwnd, EM_GETLINECOUNT, ZERO, ZERO)
    'safety check... should always be true
    If TotalLines > CurrentLine Then
        Do
            'increment current line
            CurrentLine = CurrentLine + 1
            'Get the position of the beginning of the line
            CharPos = SendMessage(txtLyrics.hwnd, EM_LINEINDEX, CurrentLine, ZERO)
            'get the length of the line
            LineLength = SendMessage(txtLyrics.hwnd, EM_LINELENGTH, CharPos, ZERO)
            'and keep looping until we get a non blank line or we get to the end
        Loop Until LineLength > 0 Or CurrentLine = TotalLines
        'if charpos = -1 then we are at the end.  Send them back to beginning
        If CharPos = -1 Then CharPos = 0
        'place cursor
        txtLyrics.SelStart = CharPos
    End If

    '/* update tags
    If Trim(txtLyrics.Text) <> "" Then Update_Tags_Ref

    'and set the focus back to the text box
    txtLyrics.SetFocus
    Exit Sub
hell:

End Sub

Private Sub vkFiletags_OLEDragDrop(Data As DataObject, Effect As Long, MouseButton As Integer, Shift As Integer, X As Single, Y As Single)
    Dim icnt As Integer
    'If Effect <> 7 And Effect <> 1 Then Exit Sub
    ''Debug.Print Data.GetData(vbCFText)
    On Error Resume Next
    If Data.Files.Count = 0 Then Exit Sub
    If Data.GetFormat(vbCFFiles) = False Then Exit Sub
    vkFiletags.UnRefreshControl = True
    For icnt = 1 To Data.Files.Count
        If FileExists(Data.Files(icnt)) Then
            If cGetInputState() <> 0 Then DoEvents
            If UCase(getFileExtensionFromPath(Data.Files(icnt))) = "MP3" Then
                Call frmTags.Load_Tags(Data.Files(icnt))
            End If

        End If
1:
    Next icnt
    vkFiletags.UnRefreshControl = False
    vkFiletags.Refresh
End Sub

Private Sub vkUpDownArt_Change(Value As Currency)
    Dim k As Long
    'On Error GoTo hell
    If MaxIndex <= 0 Or bUpdatePicArt = False Then Exit Sub
    CurIndex = Value
    'Debug.Print Value & MaxIndex & vkUpDownArt.Max
    If ReadAlbumArt(currentMp3file, CurIndex, tPic, k) Then ResizePic
hell:
End Sub


