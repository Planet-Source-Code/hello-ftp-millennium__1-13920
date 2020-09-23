VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   6975
   ClientLeft      =   5085
   ClientTop       =   930
   ClientWidth     =   3540
   LinkTopic       =   "Form3"
   ScaleHeight     =   6975
   ScaleWidth      =   3540
   Begin VB.ListBox namebox 
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox passwordbox 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox portbox 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox filelist 
      Height          =   5520
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Saved FTP Addresses"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub filelist_DblClick()

    Form1.ftpAddress.Text = filelist
    Form1.port.Text = portbox.List(filelist.ListIndex)
    Form1.usrName.Text = namebox.List(filelist.ListIndex)
    Form1.usrPassword.Text = passwordbox.List(filelist.ListIndex)
    Unload Me
End Sub
