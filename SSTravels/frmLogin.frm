VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00400000&
   Caption         =   "User Login"
   ClientHeight    =   3465
   ClientLeft      =   4125
   ClientTop       =   3240
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3465
   ScaleWidth      =   3450
   Begin VB.ListBox txtUserName 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1815
      Left            =   240
      TabIndex        =   5
      Top             =   570
      Width           =   3075
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2880
      Top             =   3210
   End
   Begin VB.TextBox txtDay 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   1620
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   1500
   End
   Begin VB.TextBox txtpassword 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "Ï"
      TabIndex        =   2
      Top             =   2520
      Width           =   3090
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   495
      Left            =   1800
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Dim xText
Dim adoPrimaryRS As New ADODB.Recordset
Private Sub cmdOK_Click()
With adoPrimaryRS
        .Requery
        .Find "[UserName] like '" & txtUserName.Text & "'"
                          
    If Not .EOF Then
    
                
        If UCase(txtpassword.Text) = UCase(.Fields("Password")) Then
                        
                If .Fields("UserLevel") = "Salesman" Then
                        
                    
                    adoPrimaryRS.Close
                    mdiMain.sb1.Panels(5).Text = frmLogin.txtUserName.Text
                    Unload Me
                    frmSales.Show
                    mdiMain.Toolbar1.Visible = False
                Else
                    adoPrimaryRS.Close
                    mdiMain.sb1.Panels(5).Text = frmLogin.txtUserName.Text
                    Unload Me
                    mdiMain.Toolbar1.Visible = True
                    mdiMain.Show
                End If
                
               
         Else
                ctr = ctr + 1
                If ctr = 4 Then
                frmInvalidPass.Show
                   End
                Else
                    If ctr = 1 Then
                        frmInvalidPass.Show
                        frmInvalidPass.Label1.Top = 160
                        frmInvalidPass.Label1.Caption = "You have 3 tries only" + vbCrLf + "  Invalid Password"
                    Else
                        If ctr = 2 Then
                        frmInvalidPass.Show
                        frmInvalidPass.Label1.Top = 160
                        frmInvalidPass.Label1.Caption = "This is your Second (2) Attempt" + vbCrLf + "  Invalid Password"
                        ElseIf ctr = 3 Then
                        frmInvalidPass.Show
                        frmInvalidPass.Label1.Top = 160
                        frmInvalidPass.Label1.Caption = "This your last Attempt" + vbCrLf + "  Invalid Password"
                        End If
                    End If
                    SendKeys "{Home}+{End}"
                End If
         End If
    
    Else
    
     ans = MsgBox("Please select User", vbOKCancel + vbCritical, "User")
    
       If ans = vbCancel Then
          End
       Else
        txtUserName.ListIndex = 0
       End If
       
    End If
    
End With

End Sub
Private Sub Command2_Click()
End
End Sub
Private Sub Form_Load()
Call Opendatabase

If adoPrimaryRS.State = 1 Then Set adoPrimaryRS = Nothing

      ' adoPrimaryRS.CursorLocation = adUseClient
       
        adoPrimaryRS.Open "SELECT * FROM Users ORDER BY UserName", Cn, adOpenDynamic, adLockPessimistic
                
               If adoPrimaryRS.RecordCount = 0 Then
                    Exit Sub
                Else
                    adoPrimaryRS.MoveFirst
                        Do While Not adoPrimaryRS.EOF
                            txtUserName.AddItem IIf(IsNull(adoPrimaryRS("UserName")), "", adoPrimaryRS("UserName"))
                               adoPrimaryRS.MoveNext
                        Loop
                End If
       
End Sub
Private Sub Timer1_Timer()
txtTime.Text = Format(Time, "hh:mm:ss")
txtDay.Text = Format(Date, "mm.dd.yyyy")
End Sub
Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdOK_Click
End Sub
Private Sub txtUserName_Click()
txtpassword.SetFocus
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtpassword.SetFocus
End Sub

