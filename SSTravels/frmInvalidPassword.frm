VERSION 5.00
Begin VB.Form frmInvalidPassword 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000A&
      BorderColor     =   &H00800000&
      BorderWidth     =   4
      Height          =   1455
      Left            =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   240
      Picture         =   "frmInvalidPassword.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid Password "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   720
      TabIndex        =   1
      Top             =   90
      Width           =   1965
   End
End
Attribute VB_Name = "frmInvalidPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendKeys "{Home}+{End}"
frmLogin.txtpassword.SetFocus
Unload Me
End Sub
