VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCustomer 
   BackColor       =   &H00400000&
   Caption         =   "Customers"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   11055
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   5055
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   10815
         Begin VB.TextBox txtAddress 
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
            Height          =   645
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   2820
            Width           =   2895
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
            Left            =   3480
            Picture         =   "frmCustomer.frx":0E42
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4320
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
            Picture         =   "frmCustomer.frx":11CC
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   4320
            Width           =   705
         End
         Begin VB.TextBox txtCustName 
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
            Top             =   1800
            Width           =   2895
         End
         Begin VB.TextBox txtCustID 
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
            Height          =   400
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtAgentCode 
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
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtNIC 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0;(0)"
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
            Height          =   435
            Left            =   1680
            MaxLength       =   13
            TabIndex        =   5
            Top             =   2280
            Width           =   2895
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
            TabIndex        =   4
            Top             =   3600
            Width           =   1935
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
            Picture         =   "frmCustomer.frx":1316
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   4320
            Width           =   735
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   840
            Width           =   1935
         End
         Begin MSComctlLib.ListView lv1 
            Height          =   4695
            Left            =   5040
            TabIndex        =   11
            Top             =   240
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   8281
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Customer ID"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Type"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Agent Code"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Customer Name"
               Object.Width           =   4410
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "NIC No."
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Address"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Contact No"
               Object.Width           =   2822
            EndProperty
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
            Index           =   1
            Left            =   600
            TabIndex        =   20
            Top             =   3000
            Width           =   690
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
            Index           =   17
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   1290
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
            Index           =   16
            Left            =   240
            TabIndex        =   16
            Top             =   1920
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer ID"
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
            TabIndex        =   15
            Top             =   480
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agent Code"
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
            TabIndex        =   14
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NIC No."
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
            TabIndex        =   13
            Top             =   2400
            Width           =   585
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
            Left            =   480
            TabIndex        =   12
            Top             =   3720
            Width           =   915
         End
      End
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Customers Profile"
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
      TabIndex        =   18
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboCustType_Click()
    If cboCustType.Text = "Sub-Agent" Then
        txtAgentCode.Locked = False
    Else
        txtAgentCode.Locked = True
    End If
End Sub

Private Sub cmdAdd_Click()
Call clearText
cmdSave.Enabled = True
cmdAdd.Enabled = False
Call autoOrderNo
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If txtCustName.Text = "" Or txtNIC.Text = "" Then
    MsgBox "Empty Fields!", vbInformation
    Exit Sub
Else
If rs.State = 1 Then
Set rs = Nothing
End If

rs.Open "select * from customer", Cn, adOpenDynamic, adLockPessimistic
    
    With rs
            .AddNew
            .Fields(0) = txtCustID.Text
            .Fields(2) = cboCustType.Text
            .Fields(1) = txtAgentCode.Text
            .Fields(3) = txtCustName.Text
            .Fields(4) = txtNIC.Text
            .Fields(5) = txtAddress.Text
            .Fields(6) = txtContact.Text
            .Update
            MsgBox "Updated successfully"
    
            'cmdSave.Enabled = True
            .Close
    End With
      addToList1
      cmdSave.Enabled = False
      cmdAdd.Enabled = True
End If
    Set rs = Nothing
End Sub

Private Sub Form_Load()
Call Opendatabase
Call autoOrderNo
cboCustType.AddItem "Sub Agent"
cboCustType.AddItem "Ordinary"

  addToList1
  cmdSave.Enabled = True
  cmdAdd.Enabled = False
End Sub
Sub clearText()
txtCustID.Text = ""
txtAgentCode.Text = ""
txtCustName.Text = ""
txtNIC.Text = ""
txtAddress.Text = ""
txtContact.Text = ""
End Sub
Sub addToList1()
lv1.ListItems.Clear
If rs.State = adStateOpen Then
        rs.Close
End If
    
    rs.Open "select * from customer where customer_id <> 0", Cn, adOpenDynamic, adLockPessimistic
    
    With rs
        If .BOF = True Or .EOF = True Then
            Exit Sub
        Else
          Do While Not .EOF
                
                Set lst = lv1.ListItems.Add(, , .Fields(0))
                    
                    lst.SubItems(1) = .Fields(1)
                    lst.SubItems(2) = .Fields(2)
                    lst.SubItems(3) = .Fields(3)
                    lst.SubItems(4) = .Fields(4)
                    lst.SubItems(5) = .Fields(5)
                    lst.SubItems(6) = .Fields(6)
                    
                    .MoveNext
            Loop
        End If
    End With
        
    Set rs = Nothing

End Sub

' Generate Customer ID
Private Sub autoOrderNo()
Dim selectRecord As New ADODB.Recordset

    selectRecord.Open "Select max(customer_id) from customer;", Cn
        If selectRecord.EOF = False Then
            txtCustID.Text = selectRecord.Fields(0) + 1
        End If
    selectRecord.Close

End Sub
