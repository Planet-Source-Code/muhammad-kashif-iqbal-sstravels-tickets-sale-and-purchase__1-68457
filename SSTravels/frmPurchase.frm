VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPurchase 
   BackColor       =   &H00400000&
   Caption         =   "Purchase"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmPurchase.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   794
      TabMaxWidth     =   3528
      TabCaption(0)   =   "Purchase Orders"
      TabPicture(0)   =   "frmPurchase.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Purchase Details"
      TabPicture(1)   =   "frmPurchase.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frame1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5535
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10575
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Height          =   5415
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   10335
            Begin VB.CommandButton cmdSave 
               BackColor       =   &H8000000A&
               Caption         =   "&Save"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   3480
               Picture         =   "frmPurchase.frx":03C2
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   4680
               Width           =   705
            End
            Begin VB.CommandButton cmdAdd 
               BackColor       =   &H8000000A&
               Caption         =   "&Add"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   2760
               Picture         =   "frmPurchase.frx":074C
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   4680
               Width           =   705
            End
            Begin VB.TextBox txtStatus 
               Appearance      =   0  'Flat
               DataField       =   "CompanyName"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   1680
               TabIndex        =   14
               Top             =   1800
               Width           =   1575
            End
            Begin VB.TextBox txtName 
               Appearance      =   0  'Flat
               DataField       =   "Address"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   1680
               TabIndex        =   13
               Top             =   1320
               Width           =   2895
            End
            Begin VB.TextBox txtACode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "CompanyID"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   1680
               TabIndex        =   12
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtNCode 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               DataField       =   "CompanyID"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   1680
               TabIndex        =   11
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox txtType 
               Appearance      =   0  'Flat
               DataField       =   "CompanyID"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   1680
               TabIndex        =   10
               Top             =   2280
               Width           =   1575
            End
            Begin VB.TextBox txtAddress 
               Appearance      =   0  'Flat
               DataField       =   "CompanyID"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   1680
               TabIndex        =   9
               Top             =   2760
               Width           =   2895
            End
            Begin VB.TextBox txtCity 
               Appearance      =   0  'Flat
               DataField       =   "Address"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   1680
               TabIndex        =   8
               Top             =   3240
               Width           =   1935
            End
            Begin VB.TextBox txtContact 
               Appearance      =   0  'Flat
               DataField       =   "Address"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   1680
               TabIndex        =   7
               Top             =   3720
               Width           =   1935
            End
            Begin VB.TextBox txtCommRate 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               DataField       =   "Address"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   1680
               TabIndex        =   6
               Top             =   4200
               Width           =   1095
            End
            Begin VB.CommandButton cmdClose 
               BackColor       =   &H8000000A&
               Caption         =   "Close"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   4200
               Picture         =   "frmPurchase.frx":0896
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   4680
               Width           =   735
            End
            Begin MSComctlLib.ListView lv1 
               Height          =   4935
               Left            =   5040
               TabIndex        =   17
               Top             =   240
               Width           =   5175
               _ExtentX        =   9128
               _ExtentY        =   8705
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               GridLines       =   -1  'True
               HotTracking     =   -1  'True
               HoverSelection  =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483629
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   9
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Alphabet Code"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Numeric Code"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Name"
                  Object.Width           =   2646
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Status"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Type"
                  Object.Width           =   2117
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   5
                  Text            =   "Address"
                  Object.Width           =   3528
               EndProperty
               BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   6
                  Text            =   "City"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   7
                  Text            =   "Contact No"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   8
                  Text            =   "Commission Rate"
                  Object.Width           =   3528
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Status"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   17
               Left            =   600
               TabIndex        =   26
               Top             =   1920
               Width           =   555
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   16
               Left            =   600
               TabIndex        =   25
               Top             =   1440
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Alphabet Code"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   15
               Left            =   240
               TabIndex        =   24
               Top             =   480
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Numeric Code"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   14
               Left            =   240
               TabIndex        =   23
               Top             =   960
               Width           =   1155
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Type"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   13
               Left            =   720
               TabIndex        =   22
               Top             =   2400
               Width           =   420
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   7
               Left            =   600
               TabIndex        =   21
               Top             =   2880
               Width           =   690
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "City"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   6
               Left            =   840
               TabIndex        =   20
               Top             =   3360
               Width           =   330
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact No"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   5
               Left            =   360
               TabIndex        =   19
               Top             =   3840
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Commission Rate"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   18
               Top             =   4320
               Width           =   1470
            End
         End
      End
      Begin VB.Frame frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5535
         Index           =   1
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   10575
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Height          =   5415
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   10335
         End
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   600
      TabIndex        =   27
      Top             =   240
      Width           =   1515
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
