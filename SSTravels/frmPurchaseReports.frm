VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPurchaseReports 
   BackColor       =   &H00400000&
   Caption         =   "Purchase Reports"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000A&
         Caption         =   "Generate Purchase Received Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3600
         Picture         =   "frmPurchaseReports.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdPurOrder 
         BackColor       =   &H8000000A&
         Caption         =   "Generate Purchase Orders Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1440
         Picture         =   "frmPurchaseReports.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker toDate 
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   44957697
         CurrentDate     =   39187
      End
      Begin MSComCtl2.DTPicker fromDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   44957697
         CurrentDate     =   39187
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   3840
         TabIndex        =   3
         Top             =   480
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
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   14
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   435
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Reports"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2865
   End
End
Attribute VB_Name = "frmPurchaseReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPurOrder_Click()
On Error Resume Next
SSTravelsDE.PurchaseOrder fromDate.Value, toDate.Value
drptSales.Sections("Section2").Controls("Label1").Caption = "Purchase Order for:  " & fromDate.Value & "  to  " & toDate.Value
drptPurchaseorder.Show
End Sub

Private Sub Form_Load()
Call Opendatabase
fromDate.Value = Date
toDate.Value = Date
End Sub
