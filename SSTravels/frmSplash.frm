VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4470
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Height          =   4035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   6240
         Top             =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "shgd dhfgdfh fhgfjh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4080
         TabIndex        =   6
         Top             =   3240
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "Ahsan Ibad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4080
         TabIndex        =   5
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "By:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3720
         TabIndex        =   4
         Top             =   2640
         Width           =   255
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   840
         TabIndex        =   2
         Top             =   1920
         Width           =   5535
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         Caption         =   "Final Project"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000018&
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   705
         Width           =   1980
      End
   End
   Begin MSComctlLib.ProgressBar ProgLoad 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   4200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'All form ontop stuff :)
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    lblProductName.Caption = App.Title
    'Centers the form.
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub


Private Sub Form_Activate()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub


Private Sub Timer1_Timer()
    
    ProgLoad.Value = ProgLoad.Value + 5
    'If the Progress Bar (ProgLoad) is 100% then your function happens.
    If ProgLoad.Value = 100 Then
        
        'Your function, can be anything. Open another form, frmMain.show... Ect.
        frmLogin.Show
        'Unloads this form
        Unload Me
    End If

End Sub


