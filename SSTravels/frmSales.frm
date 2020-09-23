VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSales 
   BackColor       =   &H00400000&
   Caption         =   "Sales"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   11415
      Begin VB.ComboBox cboCustName 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   2280
         Width           =   2175
      End
      Begin VB.ComboBox cboCustType 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1800
         Width           =   1935
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H8000000A&
         Caption         =   "&Print"
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
         Left            =   4680
         Picture         =   "frmSales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   6240
         Width           =   735
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
         Left            =   3840
         Picture         =   "frmSales.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   6240
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
         Left            =   3000
         Picture         =   "frmSales.frx":0714
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   6240
         Width           =   705
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H8000000A&
         Caption         =   "&Close"
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
         Left            =   5520
         Picture         =   "frmSales.frx":085E
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   6240
         Width           =   735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000080&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1575
         Left            =   6840
         TabIndex        =   22
         Top             =   5280
         Width           =   3735
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "CompanyID"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "##,###0"
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
            Height          =   375
            Left            =   1320
            TabIndex        =   27
            Top             =   1080
            Width           =   2175
         End
         Begin VB.TextBox txtPayments 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "CompanyID"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "##,###0"
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
            Height          =   375
            Left            =   1320
            TabIndex        =   25
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "CompanyID"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "##,###0"
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
            Height          =   375
            Left            =   1320
            TabIndex        =   23
            Top             =   120
            Width           =   2175
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
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   8
            Left            =   360
            TabIndex        =   28
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Payment"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   26
            Top             =   720
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   24
            Top             =   240
            Width           =   675
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   4080
         TabIndex        =   9
         Top             =   1440
         Width           =   3615
         Begin VB.CommandButton Command2 
            Caption         =   "Cancel"
            Height          =   495
            Left            =   2520
            TabIndex        =   21
            Top             =   3120
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "OK"
            Height          =   495
            Left            =   1560
            TabIndex        =   20
            Top             =   3120
            Width           =   855
         End
         Begin VB.TextBox txtTo 
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
            Height          =   375
            Left            =   1440
            TabIndex        =   18
            Top             =   2640
            Width           =   2055
         End
         Begin VB.TextBox txtFrom 
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
            Height          =   375
            Left            =   1440
            TabIndex        =   16
            Top             =   2160
            Width           =   2055
         End
         Begin VB.TextBox txtFlightTime 
            Appearance      =   0  'Flat
            DataField       =   "CompanyID"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "hh:mm AMPM"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   4
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
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtFlightDate 
            Appearance      =   0  'Flat
            DataField       =   "CompanyID"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
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
            Height          =   375
            Left            =   1440
            TabIndex        =   12
            Top             =   1200
            Width           =   2055
         End
         Begin VB.TextBox txtFlightNo 
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
            Height          =   375
            Left            =   1440
            TabIndex        =   10
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Flight Details"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   345
            Index           =   11
            Left            =   840
            TabIndex        =   37
            Top             =   165
            Width           =   1920
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H80000018&
            BorderWidth     =   2
            Height          =   495
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "To"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   4
            Left            =   600
            TabIndex        =   19
            Top             =   2760
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "From"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   17
            Top             =   2280
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Flight Time"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   1800
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Flight Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   1320
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Flight No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   11
            Top             =   840
            Width           =   720
         End
      End
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtInvoiceNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H000000FF&
         Height          =   400
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin MSComctlLib.ListView lv2 
         Height          =   2415
         Left            =   165
         TabIndex        =   1
         Top             =   2760
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   4260
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ticket No"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Air_A_Code"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Price"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Flight No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Flight Date"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Flight Time"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "To"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lv1 
         Height          =   2535
         Left            =   5340
         TabIndex        =   2
         Top             =   120
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4471
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Air_A_Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ticket No"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         Index           =   10
         Left            =   360
         TabIndex        =   36
         Top             =   2400
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Type"
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
         Index           =   9
         Left            =   360
         TabIndex        =   34
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2400
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Available Tickets"
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
         Left            =   360
         TabIndex        =   7
         Top             =   1440
         Width           =   1920
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
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No"
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
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
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
      Left            =   615
      TabIndex        =   38
      Top             =   360
      Width           =   885
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Private Sub cboAirline_Click()
addToList1
Label2.Caption = lv1.ListItems.Count
End Sub

Private Sub cboCustType_Click()
CustomerName_Load
End Sub

Private Sub cmdAdd_Click()
Call autoOrderNo
lv1.Refresh
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdPrint_Click()
On Error Resume Next
If (SSTravelsDE.rsSales.State <> adstateclose) Then
MsgBox "Please close the Report named Customer Bill"
Else
SSTravelsDE.Sales txtInvoiceNo.Text
drptSales.Sections("Section2").Controls("Label2").Caption = "Invoice No: " & txtInvoiceNo.Text
drptSales.Sections("Section2").Controls("Label1").Caption = "Customer Name: " & cboCustName.Text
drptSales.Sections("Section5").Controls("Label19").Caption = txtAmount.Text
drptSales.Sections("Section5").Controls("Label20").Caption = txtPayments.Text
drptSales.Sections("Section5").Controls("Label21").Caption = txtBalance.Text
'drptSales.PrintReport
drptSales.Show
End If
End Sub

Private Sub cmdSave_Click()
Dim adoOrders As New ADODB.Recordset
Dim adoInvoice As New ADODB.Recordset

If lv2.ListItems.Count <= 0 Then
        Exit Sub
Else
                        
    If txtPayments.Text = Empty Then

    MsgBox "Please enter amount!", vbInformation
        txtPayments.SetFocus
            Exit Sub
Else
    If txtInvoiceNo.Text = Empty Then
    MsgBox "No Invoice Number!", vbInformation
            txtSearch.SetFocus
            Exit Sub
    
    Else
                    
                    If adoOrders.State = 1 Then Set adoOrders = Nothing
                    
                      adoOrders.Open "SELECT * from [Sales_Invoice] where [Invoice_No] = '" & txtInvoiceNo.Text & "'", Cn, adOpenDynamic, adLockPessimistic
                    
                          With adoOrders
                                          
                              
                              If .EOF Then
                                  Cn.BeginTrans
                                  .AddNew
                                  .Fields(0) = txtInvoiceNo.Text
                                  .Fields(1) = txtAmount.Text
                                  .Fields(2) = txtPayments.Text
                                  .Fields(3) = txtBalance.Text
                                  .Fields(4) = Format(Date, "dd/mm/yyyy")
                                  .Fields(5) = Format(Time, "hh:mm")
                                  .Fields(6) = cboCustName.Text
                                  '.Fields(8) = lblLocation.Caption
                                  '.Fields(9) = mdiMain.sb1.Panels(5).Text
                                  '.Fields(10) = MonthName(Month(Date))
                                  '.Fields(11) = Year(Date)
                                  .Update
                                  .Requery
                                  Cn.CommitTrans
                                  .Close
                              End If
                          End With
                    
                    
                     
                    For i = 1 To lv2.ListItems.Count
                    
                    If adoInvoice.State = 1 Then Set adoInvoice = Nothing
                    
                      adoInvoice.Open "SELECT * from Sales_Details", Cn, adOpenDynamic, adLockPessimistic
                      
                          With adoInvoice
                          
                                  Cn.BeginTrans
                                  .AddNew
                                  .Fields(0) = txtInvoiceNo.Text
                                  .Fields(1) = lv2.ListItems(i).Text
                                  .Fields(2) = lv2.ListItems(i).SubItems(1)
                                  .Fields(3) = lv2.ListItems(i).SubItems(2)
                                  .Fields(4) = lv2.ListItems(i).SubItems(3)
                                  .Fields(5) = lv2.ListItems(i).SubItems(4)
                                  .Fields(6) = lv2.ListItems(i).SubItems(5)
                                  .Fields(7) = lv2.ListItems(i).SubItems(6)
                                  .Fields(8) = lv2.ListItems(i).SubItems(7)
                                  .Update
                                  .Requery
                                  Cn.CommitTrans
                                  .Close
                                  
                          End With
                          
                                  
                                  kashif = "update stock_details set status =  1 where [ticket_no] ='" & lv2.ListItems(i).Text & "'"
                                                     
                                                     Cn.Execute kashif
                    
                                  
                          
                    Next i
                    
                   cmdSave.Enabled = False
                    cmdPrint.Enabled = True
    End If

End If
End If

End Sub


Private Sub Command1_Click()
If txtFlightNo.Text = "" Or txtFrom.Text = "" Or txtTo.Text = "" Then
    MsgBox "Empty Fields!", vbInformation
    Exit Sub
Else
addtosalelist
    Frame4.Visible = False
End If
End Sub

Private Sub Command2_Click()
Frame4.Visible = False
End Sub

Private Sub Form_Load()
Call Opendatabase
Call autoOrderNo
Call AirLine_Load
Call CustomerType_Load
Frame4.Visible = False
End Sub
Sub AirLine_Load()
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from stock ", Cn, adOpenKeyset, adLockOptimistic
     
    With rs
            .MoveFirst
                Do While .EOF = False
                    cboAirline.AddItem ![air_a_code]
            .MoveNext
                Loop
    End With
End Sub
Sub addToList1()
lv1.ListItems.Clear
If rs.State = adStateOpen Then
        rs.Close
End If
    
    rs.Open "select * from stock_details where air_a_code = '" & cboAirline.Text & "' and status <> '1' ", Cn, adOpenDynamic, adLockPessimistic
    
    With rs
        If .BOF = True Or .EOF = True Then
        MsgBox "Stock is empty for this Airline", vbInformation
            Exit Sub
        Else
          Do While Not .EOF
                
                Set lst = lv1.ListItems.Add(, , rs![air_a_code])
                    
                    lst.SubItems(1) = rs![ticket_no]
                    lst.SubItems(2) = rs![purchase_price]
                    lst.SubItems(3) = rs![Type]
                                        
                    .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub
Private Sub lv2_DblClick()
If lv2.ListItems.Count = 0 Then
    MsgBox "No ticket to sale!", vbOKOnly + vbInformation, "Sale"
Else
    If MsgBox("Are you sure you want to sale  " & Chr(10) & Chr(10) & StrConv(lv1.SelectedItem.SubItems(1), vbUpperCase), vbYesNo + vbQuestion, "Sale Ticket") = vbYes Then
              
              lv2.ListItems.Remove lv2.SelectedItem.Index
              txtAmount.Text = CStr(Format(computeAmount, "########"))
    Else
            Exit Sub
    End If
End If

End Sub
Sub lv1_Click()
Frame4.Visible = True
txtFrom.Text = ""
txtTo.Text = ""
End Sub
' Compute Amount
Function computeAmount() As String
    Dim X As Long
    Dim total As Double

    For X = 1 To lv2.ListItems.Count
    
        total = Val(total) + lv2.ListItems(X).SubItems(2)
        
    Next X
    computeAmount = CStr(total)
End Function
Sub addtosalelist()
If rs.State = adStateOpen Then
        rs.Close
End If
    
    rs.Open "select * from stock_details where ticket_no = '" & lv1.SelectedItem.SubItems(1) & "' ", Cn, adOpenDynamic, adLockPessimistic
    
    With rs
        If .BOF = True Or .EOF = True Then
       ' MsgBox "Stock is empty for this Airline", vbInformation
            Exit Sub
        Else
          Do While Not .EOF
          
                Set lst = lv2.FindItem(rs![ticket_no], 1, , lvwPartial)
                If lst Is Nothing Then
                Set lst = lv2.ListItems.Add(, , rs![ticket_no])
                    
                    lst.SubItems(1) = rs![air_a_code]
                    lst.SubItems(2) = rs![purchase_price]
                    lst.SubItems(3) = txtFlightNo.Text
                    lst.SubItems(4) = txtFlightDate.Text
                    lst.SubItems(5) = txtFlightTime.Text
                    lst.SubItems(6) = txtFrom.Text
                    lst.SubItems(7) = txtTo.Text
                    txtAmount.Text = CStr(Format(computeAmount, "##,######0"))
                Else
                    MsgBox "Already exist", vbInformation
                End If
                    .MoveNext
            Loop
        End If
    End With
    Set rs = Nothing
End Sub
' Generate Invoice No
Private Sub autoOrderNo()
Dim selectRecord As New ADODB.Recordset

    selectRecord.Open "Select max(invoice_no) from sales_invoice;", Cn
        If selectRecord.EOF = False Then
            txtInvoiceNo.Text = selectRecord.Fields(0) + 1
        End If
    selectRecord.Close

End Sub
' entering payment
Private Sub txtPayments_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    a = CDbl(txtPayments.Text) - CDbl(txtAmount.Text)
    txtBalance.Text = CStr(Format(a, "##,###0"))
    cmdSave.SetFocus
End If
End Sub
Sub CustomerType_Load()
cboCustType.Clear
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select distinct customer_type from Customer where customer_id <> 0 ", Cn, adOpenKeyset, adLockOptimistic
     
    With rs
            .MoveFirst
                Do While .EOF = False
                    cboCustType.AddItem ![customer_type]
            .MoveNext
                Loop
    End With
End Sub
Sub CustomerName_Load()
cboCustName.Clear
If rs.State = adStateOpen Then
    rs.Close
End If

    rs.Open "select * from Customer where customer_type = '" & cboCustType.Text & "' ", Cn, adOpenKeyset, adLockOptimistic
     
    With rs
            .MoveFirst
                Do While .EOF = False
                    cboCustName.AddItem ![customer_name]
            .MoveNext
                Loop
    End With
End Sub

