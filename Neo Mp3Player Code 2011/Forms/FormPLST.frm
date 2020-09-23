VERSION 5.00
Begin VB.Form frmPLST 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   3660
   ControlBox      =   0   'False
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   244
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   4140
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8025
      Left            =   0
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   535
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   555
      TabIndex        =   15
      Top             =   30
      Width           =   8325
      Begin VB.PictureBox PicBottomLeft 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   0
         Picture         =   "Form1.frx":10197A
         ScaleHeight     =   34
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   574
         TabIndex        =   20
         Top             =   3030
         Width           =   8610
         Begin VB.PictureBox PicBottomRight 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   510
            Left            =   2280
            Picture         =   "Form1.frx":10FEB4
            ScaleHeight     =   510
            ScaleWidth      =   1380
            TabIndex        =   25
            Top             =   0
            Width           =   1380
            Begin SoftGList.Button LIST_OPTIONS 
               Height          =   120
               Left            =   735
               TabIndex        =   26
               Top             =   150
               Width           =   330
               _ExtentX        =   582
               _ExtentY        =   212
               ButtonColor     =   12632256
               MouseIcon       =   "Form1.frx":1124AE
               MousePointer    =   99
               PictureBack     =   "Form1.frx":11261C
               PictureNormal   =   "Form1.frx":11288E
               PictureDisabled =   "Form1.frx":112B00
               PictureDown     =   "Form1.frx":112D72
               PictureOver     =   "Form1.frx":112FE4
               Style           =   1
            End
            Begin VB.Image Image1 
               Height          =   225
               Left            =   1200
               MousePointer    =   15  'Size All
               Picture         =   "Form1.frx":113256
               Top             =   285
               Width           =   210
            End
         End
         Begin SoftGList.Button ADD 
            Height          =   120
            Left            =   330
            TabIndex        =   21
            Top             =   150
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   212
            ButtonColor     =   12632256
            MouseIcon       =   "Form1.frx":11352C
            MousePointer    =   99
            PictureBack     =   "Form1.frx":11369A
            PictureNormal   =   "Form1.frx":11390C
            PictureDisabled =   "Form1.frx":113B7E
            PictureDown     =   "Form1.frx":113DF0
            PictureOver     =   "Form1.frx":114062
            Style           =   1
         End
         Begin SoftGList.Button REMOVE 
            Height          =   120
            Left            =   710
            TabIndex        =   22
            Top             =   150
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   212
            ButtonColor     =   12632256
            MouseIcon       =   "Form1.frx":1142D4
            MousePointer    =   99
            PictureBack     =   "Form1.frx":114442
            PictureNormal   =   "Form1.frx":1146B4
            PictureDisabled =   "Form1.frx":114926
            PictureDown     =   "Form1.frx":114B98
            PictureOver     =   "Form1.frx":114E0A
            Style           =   1
         End
         Begin SoftGList.Button UPTRACK 
            Height          =   120
            Left            =   1110
            TabIndex        =   23
            Top             =   150
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   212
            ButtonColor     =   12632256
            MouseIcon       =   "Form1.frx":11507C
            MousePointer    =   99
            PictureBack     =   "Form1.frx":1151EA
            PictureNormal   =   "Form1.frx":11545C
            PictureDisabled =   "Form1.frx":1156CE
            PictureDown     =   "Form1.frx":115940
            PictureOver     =   "Form1.frx":115BB2
            Style           =   1
         End
         Begin SoftGList.Button DOWNTRACK 
            Height          =   120
            Left            =   1510
            TabIndex        =   24
            Top             =   150
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   212
            ButtonColor     =   12632256
            MouseIcon       =   "Form1.frx":115E24
            MousePointer    =   99
            PictureBack     =   "Form1.frx":115F92
            PictureNormal   =   "Form1.frx":116204
            PictureDisabled =   "Form1.frx":116476
            PictureDown     =   "Form1.frx":1166E8
            PictureOver     =   "Form1.frx":11695A
            Style           =   1
         End
      End
      Begin VB.PictureBox picList 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         FillColor       =   &H00F4F5F7&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2475
         Left            =   330
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   204
         TabIndex        =   16
         Top             =   570
         Width           =   3060
         Begin VB.PictureBox picBack 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6435
            Left            =   2940
            MouseIcon       =   "Form1.frx":116BCC
            MousePointer    =   99  'Custom
            Picture         =   "Form1.frx":116D2A
            ScaleHeight     =   429
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   8
            TabIndex        =   17
            Top             =   0
            Width           =   120
            Begin VB.PictureBox picBar 
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               MousePointer    =   99  'Custom
               Picture         =   "Form1.frx":119C58
               ScaleHeight     =   17
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   10
               TabIndex        =   18
               Top             =   0
               Visible         =   0   'False
               Width           =   150
            End
         End
      End
      Begin VB.PictureBox Picright 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6960
         Left            =   2670
         Picture         =   "Form1.frx":119E4A
         ScaleHeight     =   464
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   66
         TabIndex        =   27
         Top             =   0
         Width           =   990
         Begin SoftGList.Button EXIT_plst 
            Height          =   90
            Left            =   555
            TabIndex        =   28
            Top             =   0
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   159
            ButtonColor     =   12632256
            MouseIcon       =   "Form1.frx":13090C
            MousePointer    =   99
            PictureBack     =   "Form1.frx":130A7A
            PictureNormal   =   "Form1.frx":130BEC
            PictureDown     =   "Form1.frx":130D5E
            PictureOver     =   "Form1.frx":130ED0
            Style           =   1
         End
         Begin SoftGList.Button min_Plst 
            Height          =   90
            Left            =   270
            TabIndex        =   29
            Top             =   0
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   159
            ButtonColor     =   12632256
            MouseIcon       =   "Form1.frx":131042
            MousePointer    =   99
            PictureBack     =   "Form1.frx":1311B0
            PictureNormal   =   "Form1.frx":131322
            PictureDown     =   "Form1.frx":131494
            PictureOver     =   "Form1.frx":131606
            Style           =   1
         End
         Begin SoftGList.Button Button3 
            Height          =   90
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   225
            _ExtentX        =   397
            _ExtentY        =   159
            ButtonColor     =   12632256
            MouseIcon       =   "Form1.frx":131778
            MousePointer    =   99
            PictureBack     =   "Form1.frx":1318E6
            PictureNormal   =   "Form1.frx":131A58
            PictureDown     =   "Form1.frx":131BCA
            PictureOver     =   "Form1.frx":131D3C
            Style           =   1
         End
      End
   End
   Begin VB.TextBox txtAdd 
      Height          =   315
      Left            =   3390
      TabIndex        =   14
      Top             =   1740
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.VScrollBar Scrollbar 
      Height          =   3915
      Left            =   5640
      TabIndex        =   13
      Top             =   210
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   750
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4170
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.PictureBox picBack1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6435
      Left            =   6810
      Picture         =   "Form1.frx":131EAE
      ScaleHeight     =   429
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   11
      Top             =   2145
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox picBarDown 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6150
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":134DDC
      ScaleHeight     =   15.111
      ScaleMode       =   0  'User
      ScaleWidth      =   8
      TabIndex        =   10
      Top             =   870
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6120
      Top             =   4200
   End
   Begin VB.PictureBox picBarOver 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7530
      MousePointer    =   99  'Custom
      Picture         =   "Form1.frx":134FB6
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   9
      Top             =   1665
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   465
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton cmdShort 
      Caption         =   "&Short"
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton cmdRemove 
      Cancel          =   -1  'True
      Caption         =   "&Remove"
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   900
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   360
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label lblTest 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AaGgWwi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3480
      TabIndex        =   8
      Top             =   3780
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3390
      TabIndex        =   6
      Top             =   2820
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3390
      TabIndex        =   3
      Top             =   2430
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3390
      TabIndex        =   2
      Top             =   2490
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3390
      TabIndex        =   1
      Top             =   2175
      Width           =   90
   End
End
Attribute VB_Name = "frmPLST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MAX_PATH As Long = 260&
  Private Const API_FALSE As Long = 0&
  
  Private Const INVALID_HANDLE_VALUE As Long = (-1&)
  Private Const ERROR_NO_MORE_FILES As Long = 18&
 Dim lregion

  Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
  End Type
 Dim currentHeight As Long
  Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
  End Type

  Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName$, lpFindFileData As WIN32_FIND_DATA) As Long
  Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile&, lpFindFileData As WIN32_FIND_DATA) As Long
  Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile&) As Long
  Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName$) As Long


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private Type PointAPI
    x As Long
    y As Long
End Type
Dim Value As Long
Dim LargeChange As Integer
Dim smallchange As Integer

' Declarations
Dim iY As Long
Dim bDrag As Boolean
Dim iMin As Long
Dim iMax As Long
Dim iValue As Long
'Dim iSelected As Boolean
Private bMouseOver As Boolean, bMouseDown As Boolean
Private iLargeChange As Integer

Public Enum ePos
  Vertical = 0
  Horizontal = 1
End Enum

Private Enum eImg
  Normal = 0
  Down = 1
  Over = 2
End Enum

Dim checkup As Integer
'Dim checkdown As Boolean

'Hightest list item = 32768 [this is way enough](isn't it?)
Option Explicit
'Variable Declarations
Dim FONT_TyPE As Boolean
Dim drawShade As Boolean
Dim cList As New cList
Dim NormalSelection
Dim ExtendedSelection
Dim PreviousSelection
Dim ListRange                           'Range of list that will be displayed
Dim LstTextHeight
'List Settings
Const ItemSwap As Boolean = True     'Determines if the list items can be moved by dragging or not
'Color Presets
Dim CurrentTrack_Index As Integer

Dim CurrentTrack_Forecolor As Long
Dim List_Backcolor As Long 'Color of normal text
Dim SelectedText_Backcolor As Long 'Color of normal text
Dim NormalText_Forecolor As Long 'Color of normal text
Const SelectedText = &HFFFFFF           'Text color of normal selection
Const ExtendedSelectionText = &H808080  'Text color of a list item when selected by double clicking
Const SelectedBG = &HD2D8D9             'The background color of a selection

Dim check As Integer
Dim initx As Long
Dim inity As Long
'Private Enum CheckUp
 
'End Enum



Public Function GetSaveAsName(ByVal sInName$, ByVal sInitialDir$) As String
  
  Dim lpOFN As OPENFILENAME, sTemp$, nStrEnd&

  ' initialize the struct params
  With lpOFN
    .lStructSize = Len(lpOFN)
    
    ' if the 2K version of the common dialog dll is not present, subtract the byte count for the
    ' last three members of the struct
    If Is2KShell() = False Then .lStructSize = .lStructSize - 12
    
    .hWndOwner = hWnd
    
    ' tell it we want a "doc" extension.  filter strings are explained in
    ' the OPENFILENAME documentation in the MSDN
    .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & "Word Document (.doc)" & vbNullChar & "*.doc" & vbNullChar & vbNullChar
    .nFilterIndex = 2
    .lpstrFile = sInName & String$(700, 0)
    .nMaxFile = 700
    .lpstrFileTitle = String$(260, 0)
    .nMaxFileTitle = 260
    .lpstrInitialDir = sInitialDir
    .lpstrTitle = "Save File As..."
    .Flags = OFN_EXTENSIONDIFFERENT Or OFN_NOCHANGEDIR Or OFN_OVERWRITEPROMPT Or _
                    OFN_HIDEREADONLY Or OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN
    .lpstrDefExt = "doc"
    
    ' if the "Change File View" checkbox is checked, enable the hook proc and change
    ' the view before the dialog is displayed
    'If ckUseHook.Value = vbChecked Then
     ' .Flags = .Flags Or OFN_ENABLEHOOK
    '  .lpfnHook = ReturnProcAddress(AddressOf MDialogHook.DialogHookProc)
  '  End If
 ' End With
  
  If GetSaveFileName(lpOFN) Then
    sTemp = lpOFN.lpstrFile
    nStrEnd = InStr(sTemp, vbNullChar)
    If nStrEnd > 1 Then
      GetSaveAsName = Left$(sTemp, nStrEnd - 1)
    Else
      GetSaveAsName = vbNullString
    End If
  Else
    GetSaveAsName = vbNullString
  End If
  End With
End Function

Public Function GetOpenName(ByVal sInitialDir$) As String
  
  Dim lpOFN As OPENFILENAME, sTemp$, nStrEnd&

  ' initialize the struct params
  With lpOFN
    .lStructSize = Len(lpOFN)
    
    ' if the 2K version of the common dialog dll is not present, subtract the byte count for the
    ' last three members of the struct
    If Is2KShell() = False Then .lStructSize = .lStructSize - 12
    
    .hWndOwner = hWnd
    
    ' tell it we want to display all files.  filter strings are explained in
    ' the OPENFILENAME documentation in the MSDN
    .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .lpstrFile = String$(700, 0)
    .nMaxFile = 700
    .lpstrFileTitle = String$(260, 0)
    .nMaxFileTitle = 260
    .lpstrInitialDir = sInitialDir
    .lpstrTitle = "Open A File"
    .Flags = OFN_EXTENSIONDIFFERENT Or OFN_NOCHANGEDIR Or OFN_OVERWRITEPROMPT Or _
                    OFN_HIDEREADONLY Or OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN Or OFN_ALLOWMULTISELECT
    
    ' if the "Change File View" checkbox is checked, enable the hook proc and change
    ' the view before the dialog is displayed
    '///If ckUseHook.Value = vbChecked Then
    '///  .Flags = .Flags Or OFN_ENABLEHOOK
     '/// .lpfnHook = ReturnProcAddress(AddressOf MDialogHook.DialogHookProc)
    '///End If
  End With
  
  If GetOpenFileName(lpOFN) Then
    sTemp = lpOFN.lpstrFile
    nStrEnd = InStr(sTemp, vbNullChar)
    If nStrEnd > 1 Then
      GetOpenName = Left$(sTemp, nStrEnd - 1)
    Else
      GetOpenName = vbNullString
    End If
  Else
    GetOpenName = vbNullString
  End If
  
End Function
Private Function FileExists(ByVal sFileName$) As Boolean
  'checks if a file or dir exists
  'returns true if it does, returns false otherwise
  Dim hFile&, Win32FindData As WIN32_FIND_DATA
  
  sFileName = Trim$(sFileName)
  hFile = FindFirstFile(sFileName, Win32FindData)
  If (hFile <> INVALID_HANDLE_VALUE) And (hFile <> ERROR_NO_MORE_FILES) Then
    FileExists = True
  ElseIf GetFileAttributes(sFileName) <> (-1) Then
    ' FindFirstFile will not return the root dor of a drive so we check the attributes
    ' of sFileName in case it is the root
    FileExists = True
  End If
  
  Call FindClose(hFile)

End Function
Private Sub LoadFiles_from_Dir(ByVal spath$)

  Dim udtFindData As WIN32_FIND_DATA, hFileSearch&, sFileName$
  
 ' txFiles.Text = vbNullString

  If FileExists(spath) Then
    ' fix up the path
    If Right$(spath, 1) <> "\" Then spath = spath & "\"

    ' start the search for all files in the folder
    hFileSearch = FindFirstFile(spath & "*.mp3", udtFindData)

    If hFileSearch <> INVALID_HANDLE_VALUE Then
      ' get all of the file names....
      Do
        sFileName = StripNulls(udtFindData.cFileName)
                
        cList.AddItem sFileName, spath + sFileName '= txFiles.Text & sFileName
        
     '   If udtFindData.dwFileAttributes And vbDirectory Then txFiles.Text = txFiles.Text & " (Dir)"
        
        If FindNextFile(hFileSearch, udtFindData) = API_FALSE Then
          ' if we get ERROR_NO_MORE_FILES close the search and jump out of the loop
          If Err.LastDllError = ERROR_NO_MORE_FILES Then
            Call FindClose(hFileSearch)
            Exit Do
          End If
        End If
        
       ' txFiles.Text = txFiles.Text & vbNewLine
      Loop
    Else
     ' txFiles.Text = "No Files Found"
    End If
  Else
  '  txFiles.Text = "Invalid Path"
  End If
  ListRange = Fix(picList.ScaleHeight / LstTextHeight)
If cList.ItemCount > ListRange Then
    Scrollbar.Max = cList.ItemCount - ListRange 'The list items are too much to _
                                                 display we need to scroll the list
Else
    Scrollbar.Max = 0                           'There is no need for scrolling
End If
iMax = Scrollbar.Max
Scrollbar.Value = Scrollbar.Max
'ReinitializeList

End Sub
Public Function StripNulls(ByVal sText As String) As String
  ' strips any nulls from the end of a string
  Dim nPosition&
  
  StripNulls = sText
  
  nPosition = InStr(sText, vbNullChar)
  If nPosition Then StripNulls = Left$(sText, nPosition - 1)
  If Len(sText) Then If Left$(sText, 1) = vbNullChar Then StripNulls = vbNullString
End Function

Private Sub ADD_using_control_Click()

'Dim temptrack As file
Dim k
Dim listname As String
Dim dummylist As String
Dim tracklength As String
Dim kl As String
Dim spposition, pathname, trackno, i
Dim filenames As String
'cd1.DialogTitle = "Add Files"
Dim a, b As Long
On Error GoTo 2
'Me.Enabled = False
'cd1.Action = 1
'filenames = cd1.FileName
filenames = GetOpenName(App.Path)
If filenames = vbNullString Then Exit Sub
spposition = InStr(filenames, Chr(0))
If (spposition = 0) Then
'temptrack.Path = cd1.FileName
'listname = Str(totaltracks) + ". " + cd1.FileTitle
'listname = Left(listname, Len(listname) - 4)
'If Len(listname) >= 50 Then listname = Left(listname, 34) + "..."
'listname = LCase(listname)
'track.name = LCase(cd1.FileTitle)

'musicsystem.MediaPlayer2.FileName = track.path
'dummylist = listname
'listname = listname + " (" + tracklength + ")"

cList.AddItem listname, listname
'temptrack.Name = listname 'Str(totaltracks) + ". " + Left(cd1.FileTitle, 30) + Space(7) + tracklength
'i = InStr(listname, ".")
'temptrack.Name = LCase(Mid(listname, i + 1, Len(listname)))
'track(totaltracks) = temptrack
Else
    pathname = Left(filenames, spposition - 1)
    If Right(pathname, 1) <> "\" Then pathname = pathname + "\"
      filenames = Mid(filenames, spposition + 1)
' then extract each space delimited file name
    If Len(filenames) = 0 Then
       ' List1.AddItem "No files selected"
        Exit Sub
    Else
        spposition = InStr(filenames, Chr(0))
        While spposition > 0
        'totaltracks = totaltracks + 1
            'temptrack.Path = pathname + Left(filenames, spposition - 1)
            'temptrack.Name = LCase(Left(filenames, spposition - 1))
            'musicsystem.MediaPlayer2.FileName = track.path
            'tracklength = Format((Int(musicsystem.MediaPlayer2.Duration / 60)), "00") + ":" + Format(Int(Int(musicsystem.MediaPlayer2.Duration Mod 60)), "00")
            listname = Left(filenames, spposition - 1)
            'listname = Left(listname, Len(listname) - 4)

            'If TextWidth(listname) > TextWidth(Text1.Text) - 817 Then
             'If Len(listname) >= 50 Then listname = LCase(Str(totaltracks) + ". " + Left(temptrack.Name, 34)) + "..."
 
         'listname = listname + " (" + tracklength + ")"
             cList.AddItem listname, listname
              'DoEvents

             ' i = InStr(listname, ".")
             'temptrack.Name = LCase(Mid(listname, i + 1, Len(listname)))
              
              'track(totaltracks) = temptrack
              
            filenames = Mid(filenames, spposition + 1)
            spposition = InStr(filenames, Chr(0))
        Wend
cList.AddItem filenames, filenames
End If
End If
'Call cList.AddItem(txtAdd.Text, "", 0)
'Label2.Caption = "ListCount: " & cList.ItemCount
ListRange = Fix(picList.ScaleHeight / LstTextHeight)
If cList.ItemCount > ListRange Then
    Scrollbar.Max = cList.ItemCount - ListRange 'The list items are too much to _
                                                 display we need to scroll the list
Else
    Scrollbar.Max = 0                           'There is no need for scrolling
End If
iMax = Scrollbar.Max
ReinitializeList
2:
Exit Sub
End Sub

Private Sub Button1_Click()

End Sub





Private Sub ADD_MULTIPLE_FILES()
 
  Dim lpOFN As OPENFILENAME, sTemp1, nStrEnd&

  ' initialize the struct params
  With lpOFN
    .lStructSize = Len(lpOFN)
    
    ' if the 2K version of the common dialog dll is not present, subtract the byte count for the
    ' last three members of the struct
    'If Is2KShell() = False Then .lStructSize = .lStructSize - 12
    
    .hWndOwner = hWnd
    
    ' tell it we want to display all files.  filter strings are explained in
    ' the OPENFILENAME documentation in the MSDN
    .lpstrFilter = "All Files (*.*)" & vbNullChar & "*.mp3" & vbNullChar & vbNullChar
    .lpstrFile = String$(2700, 0)
    .nMaxFile = 2700
    .lpstrFileTitle = String$(1260, 0)
    .nMaxFileTitle = 1260
    '.lpstrInitialDir = sInitialDir
    .lpstrTitle = "Open A File"
    .Flags = OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_EXPLORER Or OFN_ENABLESIZING Or OFN_FORCESHOWHIDDEN Or OFN_ALLOWMULTISELECT
                    
    ' if the "Change File View" checkbox is checked, enable the hook proc and change
    ' the view before the dialog is displayed
    'If ckUseHook.Value = vbChecked Then
    ' .Flags = .Flags Or OFN_ENABLEHOOK
  '   .lpfnHook = ReturnProcAddress(AddressOf MDialogHook.DialogHookProc)
  ' End If
  End With
  
  If GetOpenFileName(lpOFN) Then
    sTemp1 = lpOFN.lpstrFile
    nStrEnd = InStr(sTemp1, vbNullChar)
    If nStrEnd > 1 Then
     ' GetOpenName = Left$(sTemp1, nStrEnd - 1)
    Else
     ' GetOpenName = vbNullString
    End If
  Else
   ' GetOpenName = vbNullString
  End If
    ' GetOpenName = lpOFN.lpstrFile
If Trim(lpOFN.lpstrFile = "") Then Exit Sub
     
     'Remove trailing null characters.
Dim nDoubleNullPos As Long
nDoubleNullPos = InStr(lpOFN.lpstrFile & vbNullChar, String$(2, 0))
Dim sfullname, spath, sExtension As String
Dim sfiletitle As String
If nDoubleNullPos Then
'Get the file name including the path name.
sfullname = Left$(lpOFN.lpstrFile, nDoubleNullPos - 1)
If sfullname = "" Then Exit Sub
'Get the file name without the path name.
sfiletitle = Left$(lpOFN.lpstrFileTitle, InStr(lpOFN.lpstrFileTitle, vbNullChar) - 1)
'Get the path name.
spath = Left$(sfullname, lpOFN.nFileOffset - 1)
'Get the extension.
If lpOFN.nFileExtension Then
sExtension = Mid$(sfullname, lpOFN.nFileExtension + 1)
End If
'If sFileTitle is a string, we have a single selection.
If Len(sfiletitle) Then
'Add to the collections.
Dim m As String
m = Mid$(sfullname, lpOFN.nFileOffset + 1)
m = m + "\" + sfullname
cList.AddItem sfiletitle, m
Else 'Tear multiple selection apart.
Dim sTemp() As String
Dim nCount As Long
sTemp = Split(sfullname, vbNullChar)
'If array contains no elements, UBound returns -1.
If UBound(sTemp) > LBound(sTemp) Then
'We have more than one array element!
'Remove backslash if sPath is the root folder.
If Len(spath) = 3 Then spath = Left$(spath, 2)
'Loop through the array, and create the
'collections; skip the first element
'(containing the path name), so start the
'counter at 1, not at 0.
For nCount = 1 To UBound(sTemp)
cList.AddItem sTemp(nCount), spath + "\" + sTemp(nCount)
'If the string already contains a backslash,
'the user must have selected a shortcut
'file, so we don't add the path.
'colFileNames.Add IIf(InStr(sTemp(nCount),sBackSlash), sTemp(nCount), sPath & sBackSlash & sTemp(nCount))
Next nCount
'Clear this variable.
'sfullname = vbNullString
End If
End If
'Add backslash if sPath is the root folder.
If Len(spath) = 2 Then
'spath = spath & sBackSlash
End If

ListRange = Fix(picList.ScaleHeight / LstTextHeight)
If cList.ItemCount > ListRange Then
    Scrollbar.Max = cList.ItemCount - ListRange 'The list items are too much to _
                                                 display we need to scroll the list
Else
    Scrollbar.Max = 0                           'There is no need for scrolling
End If
iMax = Scrollbar.Max
Scrollbar.Value = Scrollbar.Max
ReinitializeList
2:
End If





  
     
     
     
     
End Sub


Private Sub ADD_Click()
ADD_MULTIPLE_FILES
End Sub

Private Sub cmdAdd_Click()
Call cList.AddItem(txtAdd.Text, "", 0)
Label2.Caption = "ListCount: " & cList.ItemCount
ListRange = Fix(picList.ScaleHeight / LstTextHeight)
If cList.ItemCount > ListRange Then
    Scrollbar.Max = cList.ItemCount - ListRange 'The list items are too much to _
                                                 display we need to scroll the list
Else
    Scrollbar.Max = 0                           'There is no need for scrolling
End If
iMax = Scrollbar.Max
ReinitializeList

End Sub

Private Sub cmdClear_Click()
cList.Clear
ReinitializeList
End Sub

Private Sub cmdShort_Click()
cList.Sort
ReinitializeList
End Sub

Private Sub DrawBar(ImgState As eImg)
  On Error Resume Next
  Dim inty As Integer, intx As Integer
      
      
       If iMin < 0 Then iValue = -iValue Else iValue = iMax - iValue
       iY = (iValue - iMin) / (iMax - iMin) * (picBack.Height - picBar.Height)
       'inty = iY
        'If iY < picBar.Height / 2 Then iY = 0

       'If iY > picBack.ScaleHeight - (picBar.ScaleHeight / 2) Then iY = picBack.ScaleHeight - (picBar.ScaleHeight / 2)

       'If inty < picBack.Height - picBar.Height Then inty = picBack.Height - picBar.Height
         
       
       intx = 0: inty = iY
     
    
    picBack.Cls
    
    '// draw progress
         Call BitBlt(picBack.hdc, intx, inty, picBack1.ScaleWidth, picBack1.ScaleHeight, _
         picBack1.hdc, intx, inty, vbSrcCopy)
 
    '//IMAGE OVER
    If bMouseOver = True Then
       If bMouseDown = True Then
          Call BitBlt(picBack.hdc, intx, inty, picBar.ScaleWidth, picBar.ScaleHeight, _
          picBarDown.hdc, 0, 0, vbSrcCopy)
       Else
          Call BitBlt(picBack.hdc, intx, inty, picBar.ScaleWidth, picBar.ScaleHeight, _
          picBarOver.hdc, 0, 0, vbSrcCopy)
       End If
      
      picBack.Refresh
      Exit Sub
    End If

    If ImgState = Normal Then
           Call BitBlt(picBack.hdc, intx, inty, picBar.ScaleWidth, picBar.ScaleHeight, _
           picBar.hdc, 0, 0, vbSrcCopy)
    ElseIf ImgState = Down Then
           Call BitBlt(picBack.hdc, intx, inty, picBar.ScaleWidth, picBar.ScaleHeight, _
           picBarDown.hdc, 0, 0, vbSrcCopy)
    ElseIf ImgState = Over Then
           Call BitBlt(picBack.hdc, intx, inty, picBar.ScaleWidth, picBar.ScaleHeight, _
           picBarOver.hdc, 0, 0, vbSrcCopy)
    End If
        
    picBack.Refresh
End Sub

Private Sub DOWNTRACK_Click()
cList.Clear
End Sub

Private Sub EXIT_plst_Click()
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)


        '
        ' Allow moving of the form since there is no caption bar.
        '
       
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unhook
Set frmPLST = Nothing

End Sub



Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim m As PointAPI
Dim k
k = GetCursorPos(m)
initx = m.x
inity = m.y
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
Dim k
Dim m As PointAPI
k = GetCursorPos(m)
Dim k1 As Long
If Abs(m.y - inity) >= LstTextHeight Or Abs(m.x - initx) >= 2 Then
Dim lll As Integer
 lll = (picList.Height + m.y - inity) / LstTextHeight
 If lll < 1 Or lll * LstTextHeight >= 400 And (picList.Width + m.x - initx) > 500 Or (picList.Width + m.x - initx) < 204 Then Exit Sub
 'picList.Height = lll * LstTextHeight
 If lll >= 1 And lll * LstTextHeight <= 400 Then picList.Height = lll * LstTextHeight
 If (picList.Width + m.x - initx) < 500 And (picList.Width + m.x - initx) > 204 Then picList.Width = picList.Width + m.x - initx
 'picList.Width = picList.Width + m.x - initx
 picBack.Move picList.Width - picBack.Width, 0, picBack.Width, picList.Height
' picBack.Height = picList.Height
 Me.Height = (Me.Height / Me.ScaleHeight) * (picList.Height + picList.Top + PicBottomLeft.Height)
 Me.Width = (Me.Width / Me.ScaleWidth) * (picList.Width + 40)
 k = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 19, 19)
 
 Call SetWindowRgn(Me.hWnd, k, True)

 Form_Resize1
 'Me.Height = Me.Height + (m.y - inity) * 15 - 2
'Me.Width = Me.Width + (m.x - initx) * 15
' k1 = CreateRoundRectRgn(0, 0, Me.Width / 15, (Me.Height) / 15, 19, 19)

'k1 = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.Height / 15, 15, 12)
'lregion = CombineRgn(lregion, k1, lregion, RGN_OR)
'Call SetWindowRgn(Me.hWnd, k1, True)
inity = m.y
initx = m.x
'lResult = SetWindowRgn(Me.hWnd, lReigon, True)
End If
End If
End Sub


Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  iY = y
  If bDrag And Button = 1 Then '// dragging
  'iSelected = True
      '// vertical
        If y < picBar.ScaleHeight / 2 Then
         iY = 0 'iY - picBar.Height / 2
         Scrollbar.Value = Scrollbar.Min
        ElseIf y > picBack.ScaleHeight - picBar.ScaleHeight / 2 Then
         iY = picBack.ScaleHeight - picBar.ScaleHeight
         Scrollbar.Value = Scrollbar.Max
        Else
         iY = y - picBar.Height / 2
      ' If iY < picBar.ScaleHeight / 2 Then iY = picBar.ScaleHeight / 2

       'iY = y - picBar.ScaleHeight / 2
      '// horizontal
        If CalcValue <= Scrollbar.Max And CalcValue >= Scrollbar.Min Then Scrollbar.Value = Scrollbar.Max - CalcValue
  
    ' Scrollbar.Value = Scrollbar.Max - CalcValue
      ' Call DrawBar(Down)
       End If
         
  Else
    '// mouse over
           If bMouseOver = False Then
             bMouseOver = True
             'iY = y
             
           '  Call DrawBar(Over)
           End If

  End If
    
End Sub




'
'
'
'
Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim intx, inty As Integer
   If bDrag = False Then
        If CalcValue <= Scrollbar.Max And CalcValue >= Scrollbar.Min Then Scrollbar.Value = Scrollbar.Max - CalcValue
   End If
      bMouseDown = False
      iY = y
      iValue = Scrollbar.Max - Scrollbar.Value
      Call DrawBar(Normal)
      bDrag = False
       'If iMin < 0 Then iValue = -iValue Else iValue = iMax - iValue
       'iY = (iValue - iMin) / (iMax - iMin) * (picBack.Height - picBar.Height)
       'Scrollbar.Value = Scrollbar.Value
      'picBar.Picture = picBarOver.Picture
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
   

 '// vertical
    If y >= iY And y <= iY + picBar.ScaleHeight And Button = 1 Then
        If y < picBar.Height / 2 Then
        iY = 0 'iY - picBar.Height / 2
        ElseIf y > picBack.Height - picBar.Height Then
        iY = picBack.Height - picBar.Height
        Else
        iY = y - picBar.Height / 2
        End If
        bDrag = True
        bMouseDown = True
        If CalcValue = Scrollbar.Max - Scrollbar.Value Then
        Call DrawBar(Down)
        ElseIf CalcValue <= Scrollbar.Max And CalcValue >= Scrollbar.Min Then
        Scrollbar.Value = Scrollbar.Max - CalcValue
        'Call DrawBar(Normal)
        End If
        
    Else
        iY = y
        If iY > picBack.ScaleHeight - (picBar.ScaleHeight / 2) Then iY = picBack.ScaleHeight - (picBar.ScaleHeight / 2)
        If iY < picBar.ScaleHeight / 2 Then iY = picBar.ScaleHeight / 2
        iY = iY - picBar.ScaleHeight / 2
    End If
'
End If
End Sub



'
'
'
'
'
'
Private Function CalcValue() As Integer
 On Error Resume Next
    iValue = iY / (picBack.Height - picBar.Height) * (iMax - iMin) + iMin
     If iMin < 0 Then iValue = -iValue Else iValue = iMax - iValue
           'scrollbar.Value = scrollbar.Max - iValue
  CalcValue = iValue
 End Function

Private Sub Form_Load()

   Dim k
'k = SetWindowPos(Me.hWnd, -1, Me.Left / 12000 * 800, Me.Top / 9000 * 600, Me.Width / 12000 * 800, Me.Height / 9000 * 600, 0)
FONT_TyPE = True
LstTextHeight = picList.TextHeight("")
'NormalText_Forecolor = RGB(255, 0, 0)
'SelectedText_Backcolor = RGB(0, 0, 255)
'List_Backcolor = RGB(0, 255, 0)
Debug.Print LstTextHeight
Dim i
For i = 0 To Screen.FontCount - 1
cList.AddItem Screen.Fonts(i), "1234654" 'Screen.Fonts(i)
Next
CurrentTrack_Index = 4
CurrentTrack_Forecolor = RGB(255, 255, 255)
Label2.Caption = "ListCount: " & cList.ItemCount
NormalSelection = -1
ExtendedSelection = -1
 List_Backcolor = RGB(31, 50, 74)
 NormalText_Forecolor = RGB(150, 221, 255)
 SelectedText_Backcolor = RGB(17, 26, 39)
 CurrentTrack_Forecolor = RGB(78, 117, 167)
 Scrollbar.Value = Scrollbar.Min
 iValue = Scrollbar.Max
  'cd1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
currentHeight = picList.Height
Hook
Picture2.Left = 0
Picture2.Top = 0
'Me.Picture = LoadPicture(App.Path + "\" + "main1.bmp")
'Picture2.Picture = LoadPicture(App.Path + "\" + "main1.bmp")
'Picture2.AutoSize = True
'Me.Width = Picture2.Width
'Me.Height = Picture2.Height

 'lregion = fRegionFromBitmap(Picture2)
 'lregion = CombineRgn(lregion, lregion, lregion, RGN_OR)
 'lregion = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.Height / 15 - 0.7, 19, 19)

 'Call SetWindowRgn(Me.hWnd, lregion, True)
 Dim lll As Integer
 lll = (picList.Height) / LstTextHeight
 picList.Height = lll * LstTextHeight
 'picList.Width = picList.Width + m.x - initx
 picBack.Left = picList.Width - picBack.Width
 Me.Height = (Me.Height / Me.ScaleHeight) * (picList.Height + picList.Top + PicBottomLeft.Height)
 Me.Width = (Me.Width / Me.ScaleWidth) * (picList.Width + 40)
 k = CreateRoundRectRgn(0, 0, Me.ScaleWidth, Me.ScaleHeight, 19, 19)
 
 Call SetWindowRgn(Me.hWnd, k, True)

 Form_Resize1
 
 Me.Height = (Me.Height / Me.ScaleHeight) * (picList.Height + picList.Top + PicBottomLeft.Height)
 k = CreateRoundRectRgn(0, 0, Me.ScaleWidth, (Me.ScaleHeight), 19, 19)
 
 Call SetWindowRgn(Me.hWnd, k, True)

 Form_Resize1
  'Me.Picture = Picture2.Picture
  Scrollbar.Value = Scrollbar.Min
 


End Sub

Private Sub Form_Resize1()
'picList.Height = Me.ScaleHeight
'picList.Width = Me.ScaleWidth

'picList.Move 0, 0, picList.Width, Me.ScaleHeight
'picList.Move picList.Left, picList.Top, picList.ScaleWidth, Me.ScaleHeight - 2 * PicBottomLeft.Height - 4
picBack.Height = picList.Height
picBack.Move picList.ScaleWidth - picBack.Width, 0, picBack.ScaleWidth, picList.ScaleHeight
picBar.Left = picBack.Left
PicBottomRight.Move (Me.ScaleWidth - PicBottomRight.Width), 0, PicBottomRight.Width, PicBottomRight.Height
Picright.Move (Me.ScaleWidth - Picright.Width), 0, Picright.Width, Picright.Height
ListRange = Fix(picList.ScaleHeight / LstTextHeight)
If cList.ItemCount > ListRange Then
    Scrollbar.Max = cList.ItemCount - ListRange 'The list items are too much to _
                                               display we need to scroll the list
Else
    Scrollbar.Max = 0                           'There is no need for scrolling
End If
iMax = Scrollbar.Max
iValue = Scrollbar.Max - (Scrollbar.Value)

DrawBar (Normal)
PicBottomLeft.Top = Me.ScaleHeight - PicBottomLeft.Height - 1.4
ReinitializeList
End Sub

Private Sub picBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 picBar.Picture = picBarOver.Picture
End Sub

Private Sub picList_Click()
'lblTest.FontName = cList.Item(NormalSelection)
End Sub

Private Sub picList_DblClick()
'On Error Resume Next
'cList.SetEXSelection ExtendedSelection, 0
'ExtendedSelection = NormalSelection
Debug.Print NormalSelection
'cList.SetEXSelection NormalSelection, 1
'Label4.Caption = cList.Item(ExtendedSelection)
CurrentTrack_Index = NormalSelection
Text2.Text = cList.exItem(NormalSelection)
ReinitializeList

End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Fix(y / LstTextHeight) + Scrollbar.Value >= 0 And Fix(y / LstTextHeight) + Scrollbar.Value < cList.ItemCount Then
    NormalSelection = Fix(y / LstTextHeight) + Scrollbar.Value
    If cList.ItemCount > 0 Then
        Label1.Caption = NormalSelection
        Label3.Caption = cList.Item(NormalSelection) & cList.exItem(NormalSelection)
    ReinitializeList
    End If
    PreviousSelection = NormalSelection
    If CurrentTrack_Index < NormalSelection Then
    check = 1
    ElseIf CurrentTrack_Index > NormalSelection Then
    check = -1
    Else
    check = 0
    End If
    
Else
NormalSelection = -1
ReinitializeList
End If
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ItemBuffer As String
Dim EXSelectionBuffer As Integer
If Button And Fix(y / LstTextHeight) + Scrollbar.Value >= 0 And Fix(y / LstTextHeight) + Scrollbar.Value < cList.ItemCount Then
    NormalSelection = Fix(y / LstTextHeight) + Scrollbar.Value
    'If CurrentTrack_Index < NormalSelection Then check = True
   ' ElseIf CurrentTrack_Index < NormalSelection Then
    ' check = -1
    'Else
   ' check = 0
   ' End If
    
    If NormalSelection > (Scrollbar.Value + (ListRange - 1)) Then
        Scrollbar.Value = Scrollbar.Value + 1
    ElseIf NormalSelection < Scrollbar.Value Then
        Scrollbar.Value = Scrollbar.Value - 1
    End If

    If ItemSwap Then
    ItemBuffer = cList.Item(NormalSelection)
    EXSelectionBuffer = cList.EXSelection(NormalSelection)
    
    cList.ChangeItem NormalSelection, cList.Item(PreviousSelection)
    cList.SetEXSelection NormalSelection, cList.EXSelection(PreviousSelection)
    
    cList.ChangeItem PreviousSelection, ItemBuffer
    cList.SetEXSelection PreviousSelection, EXSelectionBuffer
        
    End If
    
    If cList.ItemCount > 0 Then
        Label1.Caption = NormalSelection
        Label3.Caption = cList.Item(NormalSelection) & cList.exItem(NormalSelection)
        
        
        If (check = 1) And CurrentTrack_Index >= NormalSelection Then
              check = -1
              CurrentTrack_Index = CurrentTrack_Index + 1
        ElseIf (check = -1) And CurrentTrack_Index <= NormalSelection Then
              check = 1
              CurrentTrack_Index = CurrentTrack_Index - 1
        ElseIf check = 0 Then
              CurrentTrack_Index = NormalSelection
        End If
        
        ReinitializeList
        
    End If
    PreviousSelection = NormalSelection
End If
End Sub

Private Sub picList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim filename As String
Dim extension As String
Dim tt, icnt, ipos As Integer
'tt = totaltracks
'On errror GoTo 1
If Effect <> 7 Then Exit Sub
For icnt = 1 To Data.Files.Count
     If FileExists(Data.Files(icnt)) Then
'Text1.Text = Text1.Text + Data.Files(icnt)


            'This function will add the index of this file added to the listview in order to
            'create a sequence of playback in normal sequential mode or shuffle mode
            'Form1.Maheshmp31.AddItemToPlaylist
           ipos = InStrRev(Data.Files(icnt), ".")
            extension = Mid(Data.Files(icnt), ipos + 1, Len(Data.Files(icnt)) - ipos)
       If UCase(extension) = "MP3" Or UCase(extension) = "MP2" Or UCase(extension) = "MP1" Or UCase(extension) = "WAV" Then
               ' filename = Mid(Data.Files(iCnt), Ipos + 1, Len(Data.Files(iCnt)) - Ipos) '- 4)
             ipos = InStrRev(Data.Files(icnt), "\")
             filename = LCase(Mid(Data.Files(icnt), ipos + 1, Len(Data.Files(icnt)) - ipos - 4))
             If Len(filename) >= 50 Then filename = Left(filename, 34) + "..."
            
             cList.AddItem LCase(filename), Data.Files(icnt)
             'totaltracks = totaltracks + 1
             'track(totaltracks).Name = filename
             'track(totaltracks).Path = Data.Files(icnt)
          'ElseIf UCase(extension) = "MMPL" Then
          'Dim temptrack As file

               ' Open Data.Files(icnt) For Random As #1 Len = 155
            ' Dim i As Integer
            ' i = 0
            ' Do While (1)
              '  Get #1, i + 1, temptrack
            '    'On Error GoTo 3
             '   If temptrack.Name = "" Then Exit Do
            '   cList.AddItem Str(totaltracks + 1) + ". " + temptrack.Name
            '    track(totaltracks + 1) = temptrack
            '    totaltracks = totaltracks + 1
            '    i = i + 1
            ' Loop
            ' Close #1
      '///////////////    End If
             
             
             
             
            ' File2.Path = Data.Files(icnt)
            Else
            LoadFiles_from_Dir (Data.Files(icnt))
             
            ' Text1.Text = Data.Files(icnt)
           ' Set item = ListView1.ListItems.add(ListView1.ListItems.Count + 1, , filename)
            'item.SubItems(1) = Data.Files(iCnt)
        End If
        End If
1:
Next icnt
'totaltracks = plst.List1.ListCount
'If totaltracks > tt Then
'plst.Slider1.Max = totaltracks + 1
'scrollunit = sizescroll / totaltracks
'plst.Scroll.Top = 60 + scrollunit * totaltracks
'plst.Label18.Caption = Str(currenttrack) + "/" + Str(totaltracks)
'If plst.List1.ListCount = totaltracks Then plst.List1.ListIndex = totaltracks - 1
'End If
If cList.ItemCount > ListRange Then
    Scrollbar.Max = cList.ItemCount - ListRange 'The list items are too much to _
                                                 display we need to scroll the list
Else
    Scrollbar.Max = 0                           'There is no need for scrolling
End If
iMax = Scrollbar.Max
Scrollbar.Value = Scrollbar.Max


'ReinitializeList

End Sub


Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 Call ReleaseCapture
        Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    
End Sub

Private Sub Picture3_Click()

End Sub

Private Sub REMOVE_Click()
Dim i
Dim j As String
cList.RemoveItem (NormalSelection)
If NormalSelection = cList.ItemCount Then NormalSelection = NormalSelection - 1
ListRange = Fix(picList.ScaleHeight / LstTextHeight)
If cList.ItemCount > ListRange Then
    Scrollbar.Max = cList.ItemCount - ListRange 'The list items are too much to _
                                                 display we need to scroll the list
Else
    Scrollbar.Max = 0                           'There is no need for scrolling
End If
iMax = Scrollbar.Max

ReinitializeList
Label2.Caption = "ListCount: " & cList.ItemCount
End Sub

Private Sub ScrollBar_Change()
Dim k
iValue = Scrollbar.Max - (Scrollbar.Value)
iY = iValue / (iMax - iMin) * (picBack.Height - picBar.Height) - 1
ReinitializeList

'iValue = iMax - iY / (picBack.Height - picBar.Height) * (iMax - iMin) + iMin
DrawBar (Down)

Text1.Text = Scrollbar.Value
End Sub

Private Sub ScrollBar_Scroll()
Call ScrollBar_Change
End Sub

Sub ReinitializeList()
Dim i
'Sets the scrollbar max value as the list is changed
'so the list items can be scrolled
iMax = Scrollbar.Max
picList.Cls
picList.BackColor = List_Backcolor
'Gradient
If cList.ItemCount <= ListRange Then
    For i = 0 To cList.ItemCount - 1
        If i = NormalSelection Then
            DrawSelection i
            picList.CurrentX = 0
            picList.CurrentY = NormalSelection * LstTextHeight
            
            If cList.EXSelection(i) = 1 Then
            picList.ForeColor = ExtendedSelectionText
            picList.FontBold = FONT_TyPE
            Else
              If i <> CurrentTrack_Index Then
                 picList.ForeColor = NormalText_Forecolor
              Else
                 picList.ForeColor = CurrentTrack_Forecolor
              End If
            End If
            
            picList.Print " " & i + 1 & ". " & cList.Item(i)
            picList.FontBold = FONT_TyPE
        ElseIf cList.EXSelection(i) = 1 Then
            picList.CurrentX = 0
            picList.CurrentY = i * LstTextHeight
            If i = CurrentTrack_Index Then
            picList.ForeColor = CurrentTrack_Forecolor
            Else
            picList.ForeColor = ExtendedSelectionText
            End If
            picList.FontBold = FONT_TyPE
            picList.Print " " & i + 1 & ". " & cList.Item(i)
            picList.FontBold = FONT_TyPE
            ExtendedSelection = i
        Else
            picList.CurrentX = 0
            picList.CurrentY = i * LstTextHeight
            picList.ForeColor = NormalText_Forecolor
            picList.Print " " & i + 1 & ". " & cList.Item(i)
        End If
    Next
Else
    For i = Scrollbar.Value To Scrollbar.Value + (ListRange - 1)
        If i = NormalSelection Then
            DrawSelection (i - Scrollbar.Value)
            picList.CurrentX = 0
            picList.CurrentY = (i - Scrollbar.Value) * LstTextHeight
            
            If cList.EXSelection(i) = 1 Then
            picList.ForeColor = ExtendedSelectionText
            picList.FontBold = FONT_TyPE
            Else
            If i <> CurrentTrack_Index Then
                 picList.ForeColor = NormalText_Forecolor
              Else
                 picList.ForeColor = CurrentTrack_Forecolor
              End If
            End If
            
            picList.Print " " & i + 1 & ". " & cList.Item(i)
            picList.FontBold = FONT_TyPE
        ElseIf cList.EXSelection(i) = 1 Then
            picList.CurrentX = 0
            picList.CurrentY = (i - Scrollbar.Value) * LstTextHeight
            picList.ForeColor = ExtendedSelectionText
            picList.FontBold = FONT_TyPE
            picList.Print " " & i + 1 & ". " & cList.Item(i)
            picList.FontBold = FONT_TyPE
            ExtendedSelection = i
        Else
            picList.CurrentX = 0
            picList.CurrentY = (i - Scrollbar.Value) * LstTextHeight
            If i = CurrentTrack_Index Then
            picList.ForeColor = CurrentTrack_Forecolor
            Else
            picList.ForeColor = NormalText_Forecolor
            End If
            picList.Print " " & i + 1 & ". " & cList.Item(i)
        End If
    Next
End If
End Sub

Sub DrawSelection(y)
Dim i
For i = 0 To LstTextHeight
'SelectedText_Backcolor
  picList.Line (0, i + (y * LstTextHeight))-(picList.ScaleWidth - picBar.Width, i + (y * LstTextHeight)), SelectedText_Backcolor

 'picList.Line (0, i + (y * LstTextHeight))-(picList.ScaleWidth - picBar.Width, i + (y * LstTextHeight)), RGB(((LstTextHeight - i) * 3) + 183, ((LstTextHeight - i) * 3) + 182, ((LstTextHeight - i) * 3) + 180)
Next
If drawShade Then picList.Line (0, (y * LstTextHeight))-(picList.ScaleWidth - picBar.Width, (y * LstTextHeight) + LstTextHeight), &H929EA3, B
End Sub

Sub Gradient()
Dim i
For i = 0 To picList.ScaleHeight Step LstTextHeight / 2
 picList.Line (0, i)-(picList.ScaleWidth - 18, i + LstTextHeight / 2), RGB(i / 6 + 202, i / 6 + 201, i / 6 + 200), BF
Next
End Sub

Private Sub UPTRACK_Click()
Dim txtReturn
Dim Flags As Long
Dim txtInstruction As String
txtInstruction = "Mahesh the Hero Bahut Bada"

    Dim i As Integer

    If frmOptions.optBrowse(0) Then
        Select Case frmOptions.cmbSpecial.ListIndex
            Case 0 To 11: SpecialFolder = frmOptions.cmbSpecial.ListIndex
            Case 12 To 17: SpecialFolder = frmOptions.cmbSpecial.ListIndex + 4
            Case 18 To 19: SpecialFolder = frmOptions.cmbSpecial.ListIndex + 8
            Case 20 To 22: SpecialFolder = frmOptions.cmbSpecial.ListIndex + 12
            Case Else: StartFolder = "F:\songs\03   Who Lamhe"
        End Select

        If frmOptions.chkFlags(0).Value = 1 Then Flags = Flags + BIF_USENEWUI
        If frmOptions.chkFlags(1).Value = 1 Then Flags = Flags + BIF_EDITBOX
        If frmOptions.chkFlags(2).Value = 1 Then Flags = Flags + BIF_STATUSTEXT
        If frmOptions.chkFlags(3).Value = 1 Then Flags = Flags + BIF_BROWSEINCLUDEFILES
    Else
        Flags = BIF_BROWSEFORPRINTER
        SpecialFolder = 4
    End If

      Flags = Flags + BIF_USENEWUI
       Flags = Flags + BIF_EDITBOX
        Flags = Flags + BIF_STATUSTEXT
       ' flags = flags + BIF_BROWSEINCLUDEFILES
  StartFolder = "F:\songs\03   Who Lamhe"
   txtReturn = FolderBrowse(Me.hWnd, txtInstruction, Flags)
If txtReturn <> "" Then LoadFiles_from_Dir (txtReturn)
End Sub
