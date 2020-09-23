VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   2010
   ClientTop       =   2955
   ClientWidth     =   7935
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   7935
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "FTPM File File (*.ftm)|*.ftm|"
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6360
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   6360
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   6360
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox usrPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   7
      Text            =   "no@way.com"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox usrName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "anonymous"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.TextBox port 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Text            =   "21"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox ftpAddress 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Password"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "User Name"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Port"
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
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FTP Address"
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
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
    Form1.Show
End Sub

Private Sub cmdLoad_Click()
    Dim openFile As String
    Dim temp As String
    Dim i As Integer
    
    On Error GoTo error
    Dialog1.ShowOpen                           ' opens a dialog box
        openFile = Dialog1.FileName          ' sets "openfile" to the *.pvm you choose
        Open openFile For Input As #1      ' opens that file for input to collection box

    i = 0
    Do Until EOF(1)                               ' reads the input file until the end
     Input #1, temp
     Form3.filelist.List(i) = temp              ' puts the url into the url box
     Input #1, temp
     Form3.portbox.List(i) = temp           ' puts the port into the port box
     Input #1, temp
     Form3.namebox.List(i) = temp          ' puts the user name into the name box
     Input #1, temp
     Form3.passwordbox.List(i) = temp     ' puts the password into the paswd box
     i = i + 1                                   ' moves to the next record
    Loop
    
    Close #1                                        ' closes the input file
    Form3.Show
Exit Sub

error:
     If Dialog1.CancelError = False Then Exit Sub      ' if cancel is pressed
   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub cmdOk_Click()
    Form1.Inet1.URL = ftpAddress.Text
    Form1.Inet1.RemotePort = port.Text
    Form1.Inet1.UserName = usrName.Text
    Form1.Inet1.Password = usrPassword.Text
    
    If ftpAddress.Text = "" Then
         MsgBox "You have to enter a FTP address.", vbOKOnly + vbExclamation, "Empty Collection"
       Exit Sub
    ElseIf port.Text = "" Then
         MsgBox "You have to enter a port number.", vbOKOnly + vbExclamation, "Empty Collection"
       Exit Sub
    ElseIf usrName.Text = "" Then
        MsgBox "You have to enter a user name.", vbOKOnly + vbExclamation, "Empty Collection"
       Exit Sub
    ElseIf usrPassword.Text = "" Then
        MsgBox "You have to enter a password.", vbOKOnly + vbExclamation, "Empty Collection"
       Exit Sub
    End If
    
    
    Unload Me
    Form1.Show
    Form1.message.Text = "Waiting for connection. . ."
    Form1.refresh_screen
End Sub

Private Sub cmdSave_Click()
    On Error GoTo error
    Dim Directory As String

     If ftpAddress.Text = "" Then
         MsgBox "You have to enter a FTP address.", vbOKOnly + vbExclamation, "Empty Collection"
       Exit Sub
    ElseIf port.Text = "" Then
         MsgBox "You have to enter a port number.", vbOKOnly + vbExclamation, "Empty Collection"
       Exit Sub
    ElseIf usrName.Text = "" Then
        MsgBox "You have to enter a user name.", vbOKOnly + vbExclamation, "Empty Collection"
       Exit Sub
    ElseIf usrPassword.Text = "" Then
        MsgBox "You have to enter a password.", vbOKOnly + vbExclamation, "Empty Collection"
       Exit Sub
    End If

    Dialog1.ShowSave                                 ' shows the save dialog box

    Directory$ = Dialog1.FileName               ' path where it is saved is the one you chose
             
        On Error GoTo error
        
        Open Directory$ For Append As #1                ' opens the one you choose for output
              Print #1, ftpAddress.Text                     ' saves corosponding info to file you chose
              Print #1, port.Text
              Print #1, usrName.Text
              Print #1, usrPassword.Text
        Close #1                                                            ' closes the output file
    Exit Sub
error:
If Dialog1.CancelError = False Then Exit Sub
MsgBox Err.Description, vbExclamation, "Error"
End Sub

