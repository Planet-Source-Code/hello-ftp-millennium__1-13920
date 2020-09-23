VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ftp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      TabIndex        =   23
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   22
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox usrPassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "*"
      TabIndex        =   21
      Top             =   5040
      Width           =   2775
   End
   Begin VB.TextBox usrName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   19
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox port 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9480
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox ftpAddress 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton delete 
      Caption         =   "Delete"
      Height          =   855
      Left            =   3840
      Picture         =   "ftp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   6495
      ItemData        =   "ftp.frx":0544
      Left            =   240
      List            =   "ftp.frx":0546
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   840
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5640
      TabIndex        =   10
      Top             =   840
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   5640
      TabIndex        =   9
      Top             =   1320
      Width           =   2775
   End
   Begin VB.FileListBox File2 
      Height          =   3405
      Left            =   5640
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton rename 
      Caption         =   "Rename"
      Height          =   855
      Left            =   3840
      Picture         =   "ftp.frx":0548
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox message 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   4
      Top             =   7680
      Width           =   8175
   End
   Begin VB.CommandButton disconnect 
      Caption         =   "Disconnect From FTP Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton upload 
      Caption         =   "Upload"
      Height          =   855
      Left            =   3840
      Picture         =   "ftp.frx":0F02
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Download 
      Caption         =   "Download"
      Height          =   855
      Left            =   3840
      Picture         =   "ftp.frx":1344
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton connect 
      Caption         =   "Connect To FTP Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   0
      URL             =   "ftp://:0"
      RequestTimeout  =   30
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   3240
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "FTPM File File (*.ftm)|*.ftm|"
   End
   Begin VB.Label ftpInfoHlp 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11280
      TabIndex        =   25
      ToolTipText     =   "FTP Info Help"
      Top             =   6480
      Width           =   255
   End
   Begin VB.Label fileTransferHlp 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   24
      ToolTipText     =   "File Transfer Help"
      Top             =   7320
      Width           =   255
   End
   Begin VB.Label Label7 
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
      Left            =   8880
      TabIndex        =   20
      Top             =   4680
      Width           =   2415
   End
   Begin VB.Label Label6 
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
      Left            =   9000
      TabIndex        =   18
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label5 
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
      Left            =   9600
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Address"
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
      Left            =   8880
      TabIndex        =   14
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "FTP Info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8880
      TabIndex        =   13
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Your Files"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "FTP Listing"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoad_Click()
    Dim openFile As String                  ' This is to open a file for input
    Dim temp As String                      ' what is located in the open file for input
    Dim i As Integer                            ' for the file box records
    
    On Error GoTo error
    Dialog1.ShowOpen                           ' opens a dialog box
        openFile = Dialog1.FileName          ' sets "openfile" to the *.ftm you choose
        Open openFile For Input As #1      ' opens that file for input to collection box

    i = 0                                               ' inputs to the first record in the file box
    Do Until EOF(1)                               ' reads the input file until the end
     Input #1, temp
     Form3.filelist.List(i) = temp              ' puts the url into the url box
     Input #1, temp
     Form3.portbox.List(i) = temp           ' puts the port into the port box
     Input #1, temp
     Form3.namebox.List(i) = temp          ' puts the user name into the name box
     Input #1, temp
     Form3.passwordbox.List(i) = temp     ' puts the password into the paswd box
     i = i + 1                                            ' moves to the next record
    Loop
    
    Close #1                                        ' closes the input file
    Form3.Show
Exit Sub

error:
     If Dialog1.CancelError = False Then Exit Sub      ' if cancel is pressed
   MsgBox Err.Description, vbExclamation, "Error"     'otherwise display the error
End Sub

Private Sub cmdSave_Click()
    On Error GoTo error
    Dim Directory As String                         ' for the file that you chose

     If ftpAddress.Text = "" Then
         MsgBox "You have to enter a FTP address.", vbOKOnly + vbExclamation, "Invalid Information"
       Exit Sub
    ElseIf port.Text = "" Then
         MsgBox "You have to enter a port number.", vbOKOnly + vbExclamation, "Invalid Information"
       Exit Sub
    ElseIf usrName.Text = "" Then
        MsgBox "You have to enter a user name.", vbOKOnly + vbExclamation, "Invalid Information"
       Exit Sub
    ElseIf usrPassword.Text = "" Then
        MsgBox "You have to enter a password.", vbOKOnly + vbExclamation, "Invalid Information"
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
If Dialog1.CancelError = False Then Exit Sub            ' if error is pressed
MsgBox Err.Description, vbExclamation, "Error"          ' otherwise display error
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub connect_Click()
    On Error GoTo error:
        If ftpAddress.Text = "" Or port.Text = "" Or usrName.Text = "" Or usrPassword.Text = "" Then
              MsgBox "Please be sure that all FTP information is correct and try again.", vbOKOnly + vbExclamation, "Invalid Connection Information"
              Exit Sub
        End If
        
        Inet1.URL = ftpAddress.Text                     ' sets all inet properties
        Inet1.RemotePort = port.Text
        Inet1.UserName = usrName.Text
        Inet1.Password = usrPassword.Text
    
   message.Text = "Waiting for connection. . ."
    refresh_screen
Exit Sub

error:
    If Err.Number = 13 Then
        MsgBox "Please be sure that all FTP information is correct and try again.", vbOKOnly + vbExclamation, "Invalid Connection Information"
        message.Text = ""
    Exit Sub
    End If
    
    MsgBox Err.Description, vbExclamation, "Error"
    message.Text = ""
 
End Sub


Private Sub delete_Click()
Dim counter As Integer
couner = 0
On Error GoTo error
If Inet1.URL <> "" Then                             ' if there is an url present in the box
  While counter <= List1.ListCount - 1
        If List1.Selected(counter) = True Then         ' finds which file is selected in the FTP list box
            If (Left(List1.Text, 2) <> "./" Or Left(List1.Text, 3) <> "../" Or Right(List1.Text, 1) <> "/") Then         ' if it is a file (not a directory)
                Inet1.Execute , "DELETE " & List1.List(counter)          ' issues the DELETE command to ftp
                message.Text = "Deleting File: " & List1.List(counter)
            End If
        End If
        counter = counter + 1                           ' moves to next record and tries again
    Wend
     refresh_screen


Else: MsgBox "You are not connected to a FTP server!"           ' if no url present in the box
End If
    Exit Sub
    
error:
    
    Select Case Err.Number
    Case 35764
    DoEvents
    Resume
    End Select
End Sub


Private Sub Dir1_Change()
    File2.Path = Dir1               ' when you change directory on YOUR FILE side
End Sub


Private Sub disconnect_Click()
On Error GoTo error
     If Inet1.URL <> "" Then            ' if an url is present in the box
        Inet1.Execute , "CLOSE"          ' issues the QUIT command to ftp
        Inet1.URL = ""                  ' cleans out all inet information
        Inet1.UserName = ""
        Inet1.Password = ""
        message.Text = "Disconnected!"
    Else: MsgBox "You are not connected to a FTP server!"         ' if no URL is present
    End If
    Exit Sub
     
error:
    Select Case Err.Number
    Case 35764
    Inet1.Cancel
    message.Text = "Disconnected!"
    End Select
End Sub


Private Sub Download_Click()
Dim counter As Integer
Dim fileSize As String


On Error GoTo error
If Inet1.URL <> "" Then                         ' as long as url is present in box
    While counter < List1.ListCount         ' finds which one is selected in FTP list
        If List1.Selected(counter) = True Then
            If (Left(List1.Text, 2) <> "./" Or Left(List1.Text, 3) <> "../" Or Right(List1.Text, 1) <> "/") Then     ' if not a directory
                Inet1.Execute , "SIZE " & """" & List1.List(counter) & """"
                message.Text = "Retrieving file information..."
                fileSize = Inet1.GetChunk(1024)
                fileSize = List1.List(counter) & " is " & fileSize
                fileSize = fileSize & " bytes.  Do you want to continue?"
                continue = MsgBox(fileSize, vbYesNo, "Continue File Transfer")
                If continue = vbYes Then
                    message.Text = "Downloading File: " & List1.List(counter)
                    Inet1.Execute , "GET " & """" & List1.List(counter) & """" & " " & """" & Dir1.Path & "\" & List1.List(counter) & """" ' issues the GET command
                    File2.Refresh
                End If
            End If
        End If
        counter = counter + 1               ' moves to next record and tries again
    Wend
refresh_screen
File2.Refresh

Else: MsgBox "You are not connected to a FTP server!"       ' if no URL is present
End If
Exit Sub

error:

    Select Case Err.Number
    Case 35764
    DoEvents
    Resume
    End Select
End Sub


Private Sub Drive1_Change()
    Dir1.Path = Drive1
End Sub

Private Sub fileTransferHlp_DblClick()
    Dim helpstring As String
    helpstring = "File Transfer Help:" + vbLf + vbLf
    helpstring = helpstring & "     1. Enter all of the information into the ' ftp info ' boxes." + vbLf
    helpstring = helpstring & "     2. Press the ' connect to FTP server ' button." + vbLf
    helpstring = helpstring & "     3. Once files are shown in the ' ftp listing ' box, you are connected to the remote server." + vbLf + vbLf
    helpstring = helpstring & "     From this point you can do several things such as the following:" + vbLf
    helpstring = helpstring & "          1. Click ' Download '  to download the selected file from the ftp listing box ." + vbLf
    helpstring = helpstring & "          2. Click ' Upload ' to upload the selected file in your files." + vbLf
    helpstring = helpstring & "          3. Click ' Rename ' to rename the selected file in the ftp listing." + vbLf
    helpstring = helpstring & "          4. Click ' Delete ' to delete the selected file in the ftp listing." + vbLf
    helpstring = helpstring & "          5. Click ' Disconnect from FTP server ' to disconnect from the remote server." + vbLf + vbLf
    helpstring = helpstring & "     Please note that you may not have rights to do all of the" + vbLf
    helpstring = helpstring & "        above operations on every remote server." + vbLf + vbLf
    helpstring = helpstring & "     In addition, if you wish to do the above operations on more than one file," + vbLf
    helpstring = helpstring & "        you can multi-select files in the ' ftp listing ' and in ' your files '."
    MsgBox (helpstring)
End Sub

Private Sub Form_Load()
    Dim varTemp As Variant
End Sub


Public Sub refresh_screen()
On Error GoTo error
    Inet1.Execute , "DIR"                   ' issues the DIR command to ftp
    message.Text = "Connected and Ready!"
    varTemp = Inet1.GetChunk(1024)      ' pulls information from ftp server

    Dim strArray() As String
    Dim intTemp As Integer
    
    List1.Clear
    strArray = Split(CStr(varTemp), Chr(13) & Chr(10))
    List1.AddItem ("../")  ' to go one level up on non UNIX based stations
    For intTemp = 0 To UBound(strArray)
        List1.AddItem (strArray(intTemp))
   Next
Exit Sub

error:

    Select Case Err.Number
    Case 35764
    DoEvents
    Resume
    End Select
End Sub

Private Sub ftpInfoHlp_DblClick()
    Dim helpstring As String
    helpstring = "FTP Info Help:" + vbLf + vbLf
    helpstring = helpstring & "     These boxes are used to enter connection information into the program." + vbLf
    helpstring = helpstring & "     Please note that the most common (and default) FTP port is 21." + vbLf
    helpstring = helpstring & "     If no port is specified to you, try that default port number."
    MsgBox (helpstring)
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

Select Case State
    Case 1
        message.Text = "Trying to resolve host..."
    Case 2
        message.Text = "Host is resolved"
    Case 3
        message.Text = "Sending connection request..."
    Case 9
        message.Text = "Disconnecting..."
    Case 10
        message.Text = "Disconnected"
    Case 11
        message.Text = "ERROR COMMUNICATING WITH HOST!"
    End Select
End Sub

Private Sub List1_DblClick()
On Error GoTo error
If Inet1.URL <> "" Then         ' if URL is present in box
    If (Left(List1.Text, 2) = "./" Or Left(List1.Text, 3) = "../" Or Right(List1.Text, 1) = "/") Then       ' if a directory (not a file)
        ' Clicking on the directory name
        message.Text = "Switching Directories. . ."
        Inet1.Execute , "cd " & List1.Text     ' issues cd command to selected directory in FTP box
        Inet1.Execute , "DIR"   ' issues the DIR command to ftp
        refresh_screen
     End If
Else: MsgBox "You are not connected to a FTP server!"       'if no URL is present
End If
Exit Sub

error:

    Select Case Err.Number
    Case 35764
    DoEvents
    Resume
    End Select
End Sub


Private Sub rename_Click()
On Error GoTo error
If Inet1.URL <> "" Then         ' if URL is present
    Dim namer As String
    namer = InputBox("What would you like to rename the file to: ")    ' for the NEW NAME of the file
    message.Text = "Renaming " & List1 & "To " & namer
    Inet1.Execute , "RENAME " & List1 & " " & namer      ' issues the RENAME command to ftp
    refresh_screen
Else: MsgBox "You are not connected to a FTP server!"       ' if URL is not present
End If
    Exit Sub
    
error:
  
    Select Case Err.Number
    Case 35764
    DoEvents
    Resume
    End Select
End Sub

Private Sub upload_Click()
Dim counter As Integer
counter = 0

On Error GoTo error
 If Inet1.URL <> "" Then            ' if URL is present
    While counter <= File2.ListCount - 1
        If File2.Selected(counter) = True Then      ' finds what file is selected in YOUR file list box
        Dim here As String
            Inet1.Execute , "PUT " & """" & File2.Path & "\" & File2.List(counter) & """" & " " & File2.List(counter)  ' issues the PUT command to ftp
            message.Text = "Uploading File: " & File2.List(counter)
            refresh_screen
        End If
        counter = counter + 1       ' moves to next record and tries again
    Wend

    refresh_screen
 Else: MsgBox "You are not connected to a FTP server!"      ' if no URL is present
 End If
Exit Sub

error:

    Select Case Err.Number
    Case 35764
    DoEvents
    Resume
    End Select
End Sub
