VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchase 
   BackColor       =   &H00400000&
   Caption         =   "Purchase"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmPurchaseDetails.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   794
      TabMaxWidth     =   3528
      TabCaption(0)   =   "Purchase Orders"
      TabPicture(0)   =   "frmPurchaseDetails.frx":038A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Purchase Entry"
      TabPicture(1)   =   "frmPurchaseDetails.frx":03A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frame1(1)"
      Tab(1).ControlCount=   1
      Begin VB.Frame frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Height          =   4695
            Left            =   120
            TabIndex        =   10
            Top             =   0
            Width           =   10695
            Begin VB.ComboBox cboAirline 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   1320
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker OrderDate 
               Height          =   375
               Left            =   1680
               TabIndex        =   2
               Top             =   1800
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   661
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   52887553
               CurrentDate     =   39168
            End
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
               Left            =   1560
               Picture         =   "frmPurchaseDetails.frx":03C2
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   3600
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
               Left            =   840
               Picture         =   "frmPurchaseDetails.frx":074C
               Style           =   1  'Graphical
               TabIndex        =   6
               Top             =   3600
               Width           =   705
            End
            Begin VB.TextBox txtOrderID 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
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
               ForeColor       =   &H000000C0&
               Height          =   400
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtAgencyCode 
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
               TabIndex        =   1
               Top             =   840
               Width           =   1695
            End
            Begin VB.TextBox txtTType 
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
               TabIndex        =   3
               Top             =   2280
               Width           =   2055
            End
            Begin VB.TextBox txtTotTickets 
               Alignment       =   2  'Center
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
               TabIndex        =   4
               Top             =   2760
               Width           =   975
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
               Left            =   2280
               Picture         =   "frmPurchaseDetails.frx":0896
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   3600
               Width           =   735
            End
            Begin MSComctlLib.ListView lv1 
               Height          =   4215
               Left            =   3840
               TabIndex        =   13
               Top             =   240
               Width           =   6735
               _ExtentX        =   11880
               _ExtentY        =   7435
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
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   6
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Order ID"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Agency Code"
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   2
                  Text            =   "Airline"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   3
                  Text            =   "Order Date"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   4
                  Text            =   "Ticket Type"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   5
                  Text            =   "Total Tickets"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Airline"
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
               TabIndex        =   19
               Top             =   1440
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Order ID"
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
               Left            =   480
               TabIndex        =   18
               Top             =   480
               Width           =   720
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Agency Code"
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
               Left            =   360
               TabIndex        =   17
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Order Date"
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
               Left            =   480
               TabIndex        =   16
               Top             =   1920
               Width           =   930
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Ticket Type"
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
               Left            =   480
               TabIndex        =   15
               Top             =   2400
               Width           =   990
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Tickets"
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
               Left            =   480
               TabIndex        =   14
               Top             =   2880
               Width           =   1095
            End
         End
      End
      Begin VB.Frame frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Index           =   1
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Height          =   4695
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   5295
            Begin VB.TextBox txtBalance 
               Alignment       =   2  'Center
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
               ForeColor       =   &H000000C0&
               Height          =   400
               Left            =   3960
               TabIndex        =   33
               Top             =   1200
               Width           =   975
            End
            Begin VB.TextBox txtAir 
               Appearance      =   0  'Flat
               DataField       =   "Address"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
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
               Left            =   1920
               TabIndex        =   32
               Top             =   720
               Width           =   2175
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Ticket Details"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   2895
               Left            =   120
               TabIndex        =   26
               Top             =   1680
               Width           =   5055
               Begin VB.TextBox txtUnitPrice 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  DataField       =   "Address"
                  BeginProperty DataFormat 
                     Type            =   0
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
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
                  TabIndex        =   38
                  Top             =   1200
                  Width           =   2415
               End
               Begin VB.CommandButton Command3 
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
                  Left            =   3480
                  Picture         =   "frmPurchaseDetails.frx":09E0
                  Style           =   1  'Graphical
                  TabIndex        =   37
                  Top             =   1920
                  Width           =   735
               End
               Begin VB.CommandButton cmdAdd2 
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
                  Left            =   2040
                  Picture         =   "frmPurchaseDetails.frx":0B2A
                  Style           =   1  'Graphical
                  TabIndex        =   36
                  Top             =   1920
                  Width           =   705
               End
               Begin VB.CommandButton cmdSave2 
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
                  Left            =   2760
                  Picture         =   "frmPurchaseDetails.frx":0C74
                  Style           =   1  'Graphical
                  TabIndex        =   35
                  Top             =   1920
                  Width           =   705
               End
               Begin VB.ComboBox cboACode 
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1680
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.TextBox txtTicketNo 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  DataField       =   "Address"
                  BeginProperty DataFormat 
                     Type            =   0
                     Format          =   "0"
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1033
                     SubFormatType   =   0
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
                  TabIndex        =   27
                  Top             =   720
                  Width           =   2415
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Unit Price"
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
                  Index           =   8
                  Left            =   600
                  TabIndex        =   39
                  Top             =   1320
                  Width           =   810
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Air A Code"
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
                  Index           =   3
                  Left            =   600
                  TabIndex        =   30
                  Top             =   360
                  Width           =   870
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Ticket No."
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
                  Index           =   2
                  Left            =   480
                  TabIndex        =   28
                  Top             =   840
                  Width           =   825
               End
            End
            Begin VB.TextBox txtTicReceived 
               Alignment       =   2  'Center
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
               Left            =   1920
               TabIndex        =   24
               Top             =   1200
               Width           =   975
            End
            Begin VB.ComboBox cboOrder 
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
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   240
               Width           =   1695
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Balance"
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
               Left            =   3120
               TabIndex        =   34
               Top             =   1320
               Width           =   660
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Airline"
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
               Left            =   840
               TabIndex        =   31
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tickets Received"
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
               Index           =   1
               Left            =   360
               TabIndex        =   25
               Top             =   1320
               Width           =   1440
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Order ID"
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
               Index           =   0
               Left            =   720
               TabIndex        =   23
               Top             =   360
               Width           =   720
            End
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
      Left            =   840
      TabIndex        =   20
      Top             =   360
      Width           =   1515
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs1 As New ADODB.Recordset

Private Sub cboACode_Click()
updateStock
cboACode.Enabled = False
txtTicReceived.Enabled = False
txtBalance.Enabled = False
cboOrder.Enabled = False
txtAir.Enabled = False
End Sub

Private Sub cboAirline_Click()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from AIRLINE where air_name = '" & cboAirline.Text & "'", Cn, adOpenKeyset, adLockOptimistic
     
    With rs
            If rs.Fields(1) = "Direct" Then
                txtTType.Text = "Stock Performa"
            Else
                txtTType.Text = "Exchange Order"
            End If
    End With
End Sub

Private Sub cboOrder_Click()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from purchase_order, purchase_details where purchase_order.order_id = '" & cboOrder.Text & "' and purchase_order.order_id = purchase_details.order_id", Cn, adOpenKeyset, adLockOptimistic
     
    With rs
            If Not .EOF Then
                txtAir.Text = rs![air_a_code]
                txtBalance.Text = rs![balance]
            End If
    End With
ACode_Load
End Sub

Private Sub cmdAdd_Click()
autoOrderNo
txtAgencyCode.Text = ""
txtTType.Text = ""
txtTotTickets.Text = ""
cmdSave.Enabled = True
cmdAdd.Enabled = False
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If txtAgencyCode.Text = "" Then
    MsgBox "Empty Fields!", vbInformation
Else
If rs.State = 1 Then
Set rs = Nothing
End If

    rs.Open "select * from PURCHASE_ORDER", Cn, adOpenDynamic, adLockPessimistic
    
    With rs
            .AddNew
            .Fields(0) = txtOrderID.Text
            .Fields(1) = txtAgencyCode.Text
            .Fields(3) = cboAirline.Text
            .Fields(4) = OrderDate.Value
    
            .Update
            MsgBox "Updated successfully"
    
            .Close
    End With
    
If rs1.State = 1 Then
Set rs1 = Nothing
End If

    rs1.Open "select * from PURCHASE_DETAILS", Cn, adOpenDynamic, adLockPessimistic
    
    With rs1
            .AddNew
            .Fields(0) = txtOrderID.Text
            .Fields(1) = txtTType.Text
            .Fields(2) = Val(txtTotTickets.Text)
            .Fields(3) = Val(txtTotTickets.Text)
    
            .Update
            MsgBox "Updated successfully"
    
            .Close
    End With
      addToList1
      cmdSave.Enabled = False
      cmdAdd.Enabled = True
End If
    Set rs = Nothing
End Sub

Private Sub cmdAdd2_Click()
txtTicketNo.Text = ""
txtUnitPrice.Text = ""
End Sub

Private Sub cmdSave2_Click()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from stock_details", Cn, adOpenKeyset, adLockOptimistic
    
    With rs
            .AddNew
            .Fields(0) = cboACode.Text
            .Fields(1) = Val(txtTicketNo.Text)
            .Fields(2) = Val(txtUnitPrice.Text)
            .Update
            MsgBox "Saved Successfully"
            .Close
    End With
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Opendatabase
OrderDate.Value = Date
autoOrderNo
AirLine_Load
OrderID_Load
addToList1
cmdSave.Enabled = True
cmdAdd.Enabled = False
End Sub
' Generate Order Number
Private Sub autoOrderNo()
Dim selectRecord As New ADODB.Recordset

    selectRecord.Open "Select max(order_id) from PURCHASE_ORDER;", Cn
        If selectRecord.EOF = False Then
            txtOrderID.Text = selectRecord.Fields(0) + 1
        End If
    selectRecord.Close

End Sub
Sub AirLine_Load()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select distinct air_name from AIRLINE ", Cn, adOpenKeyset, adLockOptimistic
     
    With rs
            .MoveFirst
                Do While .EOF = False
                    cboAirline.AddItem ![air_name]
            .MoveNext
                Loop
    End With
End Sub
Sub OrderID_Load()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from purchase_order where order_status = '0' ", Cn, adOpenKeyset, adLockOptimistic
     
    With rs
            If .EOF = True Or .BOF = True Then
            Exit Sub
            Else
            .MoveFirst
            If .EOF = True Or .BOF = True Then
            Exit Sub
            Else
            
                Do While .EOF = False
                    cboOrder.AddItem ![order_id]
            .MoveNext
                Loop
            End If
            End If
    End With
End Sub
Sub ACode_Load()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from airline where air_name = '" & txtAir.Text & "' ", Cn, adOpenKeyset, adLockOptimistic
     
    With rs
            .MoveFirst
                Do While .EOF = False
                    cboACode.AddItem ![air_a_code]
            .MoveNext
                Loop
    End With
End Sub
Sub addToList1()
lv1.ListItems.Clear
If rs.State = adStateOpen Then
        rs.Close
End If
    
    rs.Open "select * from PURCHASE_ORDER, PURCHASE_DETAILS where PURCHASE_ORDER.order_id = PURCHASE_DETAILS.order_id ", Cn, adOpenDynamic, adLockPessimistic
    
    With rs
        If .BOF = True Or .EOF = True Then
            Exit Sub
        Else
          Do While Not .EOF
                
                Set lst = lv1.ListItems.Add(, , rs![order_id])
                    
                    lst.SubItems(1) = rs![agency_code]
                    lst.SubItems(2) = rs![air_a_code]
                    lst.SubItems(3) = rs![date_placement]
                    lst.SubItems(4) = rs![ticket_type]
                    lst.SubItems(5) = rs![tickets_req]
                    
                    .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub
Private Sub txtTicReceived_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call calculateBalance
updateOrderStatus
End If
End Sub
Sub calculateBalance()
If rs.State = adStateOpen Then
        rs.Close
End If
    
    rs.Open "select * from PURCHASE_DETAILS where ORDER_ID = '" & cboOrder.Text & "' ", Cn, adOpenDynamic, adLockPessimistic
    
    With rs
            rs![tickets_recv] = Val(txtTicReceived.Text)
            txtBalance.Text = Val(txtBalance.Text) - Val(txtTicReceived.Text)
            rs![balance] = Val(txtBalance.Text)
            .Update
    End With
    Set rs = Nothing
End Sub
Sub updateOrderStatus()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from purchase_order, purchase_details where purchase_order.order_id = purchase_details.order_id and purchase_order.order_status = '0' and purchase_order.order_id = '" & cboOrder.Text & "' ", Cn, adOpenKeyset, adLockOptimistic
     
     With rs
    If .BOF = True Or .EOF = True Then
            Exit Sub
        Else
            .Update
            If rs![balance] <> 0 Then
                rs![order_status] = 0
            Else
                rs![order_status] = 1
            End If
            .Update
            .Close
       End If
    End With
    Set rs = Nothing
End Sub
Sub updateStock()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from stock where air_a_code = '" & cboACode.Text & "'", Cn, adOpenKeyset, adLockOptimistic
    
    With rs
            .Update
            .Fields(1) = .Fields(1) + (txtTicReceived.Text)
            .Update
            .Close
    End With
    Set rs = Nothing
End Sub









