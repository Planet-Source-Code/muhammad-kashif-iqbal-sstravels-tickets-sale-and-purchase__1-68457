VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStock 
   BackColor       =   &H00400000&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   600
      TabIndex        =   1
      Top             =   960
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   794
      TabMaxWidth     =   3528
      TabCaption(0)   =   "By Airline"
      TabPicture(0)   =   "frmStock.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frame1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Details"
      TabPicture(1)   =   "frmStock.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frame1(1)"
      Tab(1).ControlCount=   1
      Begin VB.Frame frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Index           =   1
         Left            =   -74880
         TabIndex        =   5
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Height          =   4695
            Left            =   120
            TabIndex        =   6
            Top             =   0
            Width           =   5295
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
               TabIndex        =   20
               Top             =   240
               Width           =   1695
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
               TabIndex        =   19
               Top             =   1200
               Width           =   975
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
               TabIndex        =   9
               Top             =   1680
               Width           =   5055
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
                  TabIndex        =   15
                  Top             =   720
                  Width           =   2415
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
                  TabIndex        =   14
                  Top             =   240
                  Width           =   1695
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
                  Picture         =   "frmStock.frx":0038
                  Style           =   1  'Graphical
                  TabIndex        =   13
                  Top             =   1920
                  Width           =   705
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
                  Picture         =   "frmStock.frx":03C2
                  Style           =   1  'Graphical
                  TabIndex        =   12
                  Top             =   1920
                  Width           =   705
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
                  Picture         =   "frmStock.frx":050C
                  Style           =   1  'Graphical
                  TabIndex        =   11
                  Top             =   1920
                  Width           =   735
               End
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
                  TabIndex        =   10
                  Top             =   1200
                  Width           =   2415
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
                  TabIndex        =   18
                  Top             =   840
                  Width           =   825
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
                  TabIndex        =   17
                  Top             =   360
                  Width           =   870
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
                  TabIndex        =   16
                  Top             =   1320
                  Width           =   810
               End
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
               TabIndex        =   8
               Top             =   720
               Width           =   2175
            End
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
               TabIndex        =   7
               Top             =   1200
               Width           =   975
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
               TabIndex        =   24
               Top             =   360
               Width           =   720
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
               TabIndex        =   23
               Top             =   1320
               Width           =   1440
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
               TabIndex        =   22
               Top             =   840
               Width           =   540
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
               TabIndex        =   21
               Top             =   1320
               Width           =   660
            End
         End
      End
      Begin VB.Frame frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4815
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   10935
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            Height          =   4695
            Left            =   120
            TabIndex        =   3
            Top             =   0
            Width           =   10695
            Begin MSComctlLib.ListView lv1 
               Height          =   4215
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   4095
               _ExtentX        =   7223
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
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Airline"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   1
                  SubItemIndex    =   1
                  Text            =   "Total Tickets"
                  Object.Width           =   2540
               EndProperty
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
      Caption         =   "Stock"
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
      Left            =   1260
      TabIndex        =   0
      Top             =   480
      Width           =   915
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call Opendatabase
addToList1
End Sub
Sub addToList1()
lv1.ListItems.Clear
If rs.State = adStateOpen Then
        rs.Close
End If
kashif = "select * from stock"
    rs.Open kashif, Cn, adOpenDynamic, adLockPessimistic
    
    With rs
        If .BOF = True Or .EOF = True Then
            Exit Sub
        Else
          Do While Not .EOF
                
                Set lst = lv1.ListItems.Add(, , rs![air_a_code])
                    
                    lst.SubItems(1) = rs![total_tickets]
                                       
                    .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub
