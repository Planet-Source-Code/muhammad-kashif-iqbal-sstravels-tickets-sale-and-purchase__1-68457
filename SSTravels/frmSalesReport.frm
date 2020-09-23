VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesReport 
   BackColor       =   &H00400000&
   Caption         =   "Sales Reports"
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
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   6735
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   2040
         ScaleHeight     =   855
         ScaleWidth      =   2895
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   2895
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
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter Invoice No"
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
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1380
         End
      End
      Begin VB.CommandButton cmdPurOrder 
         BackColor       =   &H8000000A&
         Caption         =   "Print Customer Bill"
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
         Left            =   600
         Picture         =   "frmSalesReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker toDate 
         Height          =   375
         Left            =   4200
         TabIndex        =   2
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
         Format          =   19529729
         CurrentDate     =   39187
      End
      Begin MSComCtl2.DTPicker fromDate 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
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
         Format          =   19529729
         CurrentDate     =   39187
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
         TabIndex        =   5
         Top             =   480
         Width           =   435
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
         TabIndex        =   4
         Top             =   480
         Width           =   210
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Reports"
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
      Left            =   960
      TabIndex        =   6
      Top             =   480
      Width           =   2235
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPurOrder_Click()
Picture1.Visible = True
End Sub

Private Sub Form_Load()
Call Opendatabase
fromDate.Value = Date
toDate.Value = Date
End Sub
' entering invoice no
Private Sub txtInvoiceNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    Picture1.Visible = False
    
If rs.State = adStateOpen Then
rs.Close
End If

rs.Open "select * from sales_invoice where invoice_no = ' " & txtInvoiceNo.Text & " '", Cn, adOpenDynamic, adLockPessimistic

SSTravelsDE.Sales txtInvoiceNo.Text
drptSales.Sections("Section2").Controls("Label2").Caption = "Invoice No: " & txtInvoiceNo.Text
drptSales.Sections("Section2").Controls("Label1").Caption = "Customer Name: " & rs![customer_name]
drptSales.Sections("Section5").Controls("Label19").Caption = rs![amount]
drptSales.Sections("Section5").Controls("Label20").Caption = rs![payments]
drptSales.Sections("Section5").Controls("Label21").Caption = rs![Change]

drptSales.Show

Set rs = Nothing

End If
End Sub
