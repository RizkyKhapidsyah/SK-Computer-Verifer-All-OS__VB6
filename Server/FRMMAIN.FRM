VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   Caption         =   "CS Computer Verifier - Windows 95/98/ME/NT/2000 Server"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   10650
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3240
      Top             =   2160
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   285
      Left            =   9120
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Listen for connections at startup"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CheckBox chkListen 
      Caption         =   "Listen For Connections"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Default Port"
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock TcpServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox TextPort 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "1979"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox TextIP 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear Alerts"
      Height          =   285
      Left            =   9120
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Alert Date"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   14711
      EndProperty
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   7800
      TabIndex        =   12
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Double Click to View Alert"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Listen on port:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Current IP"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Alerts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkListen_Click()
On Error GoTo tcperror
Select Case chkListen.Value
 Case "1"
  chkListen.Caption = "Disconnect"
  TextPort.Enabled = False
  TcpServer.LocalPort = TextPort.Text
  TcpServer.Listen
 Case "0"
 chkListen.Caption = "Listen For Connections"
 TextPort.Enabled = True
  TcpServer.Close
End Select
Exit Sub
tcperror:
MsgBox Error & " " & TextPort.Text, vbCritical
chkListen.Value = 0
End Sub

Private Sub Command1_Click()
On Error Resume Next
TextPort.Text = "1979"
End Sub

Private Sub Command2_Click()
Call TcpServer.Close
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\Settings.ini" For Output Access Write As #filehandle%
    Print #filehandle%, TextPort.Text
    Print #filehandle%, Check1.Value
    Close

End
End Sub

Private Sub Command3_Click()
On Error Resume Next
ListView1.ListItems.Clear
End Sub

Private Sub Form_Load()
On Error Resume Next
TextIP.Text = TcpServer.LocalIP
Dim intfile As Integer
Dim pass As String
Dim pass2 As String
intfile = FreeFile
  Open App.Path & "\Settings.ini" For Input As #intfile
  Input #intfile, pass
  Input #intfile, pass2
  TextPort.Text = pass
  Check1.Value = pass2
  Close #intfile
  
  If Check1.Value = 1 Then
  chkListen.Value = 1
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Call TcpServer.Close
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\Settings.ini" For Output Access Write As #filehandle%
    Print #filehandle%, TextPort.Text
    Print #filehandle%, Check1.Value
    Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call TcpServer.Close
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\Settings.ini" For Output Access Write As #filehandle%
    Print #filehandle%, TextPort.Text
    Print #filehandle%, Check1.Value
    Close

End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
    Dim f As frmreadalert
    Set f = New frmreadalert
    f.Text2.Text = ListView1.SelectedItem
f.Text1.Text = ListView1.SelectedItem.SubItems(1)
f.Show

End Sub
Private Sub tcpServer_Close()
Call TcpServer.Close
TcpServer.Listen 'listen for another connection
End Sub

Private Sub tcpServer_ConnectionRequest(ByVal requestID As Long)
'close the server if it is open before accepting
'the connection request
If TcpServer.State <> sckClosed Then
 TcpServer.Close
End If
 Call TcpServer.Accept(requestID) 'accepts connection
End Sub

Private Sub tcpServer_DataArrival(ByVal bytesTotal As Long)
Dim strMessage As String
'converts the data into a string called strMessage$
Call TcpServer.GetData(strMessage$)
    Set Item = ListView1.ListItems.Add(, , Format(Date, "mm-dd-yyyy") & " at " & (Format(Now, "h:mm:ss AM/PM")))
    Item.SubItems(1) = strMessage$
End Sub

Private Sub tcpServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   Dim result As Integer
   result = MsgBox(Source & ": " & Description, _
      vbOKOnly, "TCP/IP Error")
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Label5.Caption = "Total Alerts: " & ListView1.ListItems.Count
End Sub
