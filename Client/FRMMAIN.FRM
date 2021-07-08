VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmmain 
   Caption         =   "CS Computer Verifier - Windows 95/98/ME/NT/2000 Client"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6270
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   3120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   0
      Top             =   2640
   End
   Begin VB.TextBox TextServer 
      Height          =   285
      Left            =   75
      TabIndex        =   28
      Text            =   "localhost"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TextPort 
      Height          =   285
      Left            =   75
      TabIndex        =   27
      Text            =   "1979"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock TcpClient 
      Left            =   75
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Exit when done."
      Height          =   375
      Left            =   1260
      TabIndex        =   26
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Automaticly get information at program startup."
      Height          =   375
      Left            =   1260
      TabIndex        =   25
      Top             =   4560
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stored Information"
      Height          =   3975
      Left            =   555
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox Text12 
         DataField       =   "Adapter Info"
         Height          =   765
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox Text11 
         DataField       =   "Drives"
         Height          =   765
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox Text10 
         DataField       =   "Memory"
         Height          =   525
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         DataField       =   "Processor Info"
         Height          =   525
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         DataField       =   "Os"
         Height          =   765
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Adapter Info:"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   24
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Drives:"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Memory:"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Processor Info:"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Os:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Information"
      Height          =   3975
      Left            =   660
      TabIndex        =   3
      Top             =   0
      Width           =   4935
      Begin VB.TextBox text6 
         DataField       =   "Adapter Info"
         Height          =   765
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox text5 
         DataField       =   "Drives"
         Height          =   765
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox text4 
         DataField       =   "Memory"
         Height          =   525
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox text3 
         DataField       =   "Processor Info"
         Height          =   525
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox text2 
         DataField       =   "Os"
         Height          =   765
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Adapter Info:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Drives:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Memory:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Processor Info:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Os:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4875
      TabIndex        =   1
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   4875
      TabIndex        =   2
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Begin VB.Menu menusettings 
         Caption         =   "Settings"
         Shortcut        =   {F1}
      End
      Begin VB.Menu menu1 
         Caption         =   "-"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents sink As SWbemSink
Attribute sink.VB_VarHelpID = -1
Function FileExists(strFile As String) As Integer
Dim lSize As Long

On Error Resume Next
    '* set lSize to -1
lSize = -1
    'Get the length of the file
lSize = FileLen(strFile)
If lSize = 0 Then
    '* File is zero bytes and exists
FileExists = 0
ElseIf lSize > 3 Then
    '* File Exists
FileExists = 1
Else
    '* Does not exist
FileExists = -1
End If
End Function
Public Sub LoadStoredData()
On Error Resume Next
Text1.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
Dim intfile As Integer
Dim pass As String
ENTER = Chr$(13) + Chr$(10)
intfile = FreeFile
  Open App.Path & "\OS.dat" For Input As #intfile
  Do
  Input #intfile, pass
  Text1.Text = Text1.Text & pass & ENTER
  Loop Until EOF(intfile)
  Close #intfile

  Open App.Path & "\Processor.dat" For Input As #intfile
  Do
  Input #intfile, pass
  Text9.Text = Text9.Text & pass & ENTER
  Loop Until EOF(intfile)
  Close #intfile
  
  Open App.Path & "\Memory.dat" For Input As #intfile
  Do
  Input #intfile, pass
  Text10.Text = Text10.Text & pass & ENTER
  Loop Until EOF(intfile)
  Close #intfile

  Open App.Path & "\Drives.dat" For Input As #intfile
  Do
  Input #intfile, pass
  Text11.Text = Text11.Text & pass & ENTER
  Loop Until EOF(intfile)
  Close #intfile
  
  Open App.Path & "\Adapter.dat" For Input As #intfile
  Do
  Input #intfile, pass
  Text12.Text = Text12.Text & pass & ENTER
  Loop Until EOF(intfile)
  Close #intfile
End Sub
Public Sub CompareData()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
' set up local port and wait for connection
TcpClient.RemoteHost = TextServer.Text
   
If TcpClient.RemoteHost = "" Then
      TcpClient.RemoteHost = "localhost"
End If
   
TcpClient.RemotePort = TextPort.Text  ' server port
Call TcpClient.Connect  ' connect to RemoteHost address
DoEvents
DoEvents
DoEvents

If Text1.Text = "" Then
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "There is nothing stored for Operating System Information")
Else
If text2.Text = Text1.Text Then
Else
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "Operating System information doesn't match the stored information." & ENTER & ENTER & "Information Stored Is:" & ENTER & Text1.Text & ENTER & "Information Found Was:" & ENTER & text2.Text)
End If
End If
DoEvents
If Text9.Text = "" Then
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "There is nothing stored for Processor Information")
Else
If text3.Text = Text9.Text Then
Else
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "Processor information doesn't match the stored information." & ENTER & ENTER & "Information Stored Is:" & ENTER & Text9.Text & ENTER & "Information Found Was:" & ENTER & text3.Text)
End If
End If
DoEvents
If Text10.Text = "" Then
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "There is nothing stored for Memory Information.")
Else
If text4.Text = Text10.Text Then
Else
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "Memory information doesn't match the stored information." & ENTER & ENTER & "Information Stored Is:" & ENTER & Text10.Text & ENTER & "Information Found Was:" & ENTER & text4.Text)
End If
End If
DoEvents
If Text11.Text = "" Then
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "There is nothing stored for Drive Information.")
Else
If text5.Text = Text11.Text Then
Else
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "Drive information doesn't match the stored information." & ENTER & ENTER & "Information Stored Is:" & ENTER & Text11.Text & ENTER & "Information Found Was:" & ENTER & text5.Text)
End If
End If
DoEvents
If Text12.Text = "" Then
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "There is nothing stored for Adapter Information.")
Else
If text6.Text = Text12.Text Then
Else
Call TcpClient.SendData("Computer " & TcpClient.LocalHostName & ":" & ENTER & ENTER & "Adapter information doesn't match the stored information." & ENTER & ENTER & "Information Stored Is:" & ENTER & Text12.Text & ENTER & "Information Found Was:" & ENTER & text6.Text)
End If
End If
DoEvents
TcpClient.Close
End Sub

Public Sub GetOSData()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
'Os Information
Set SystemSet = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")

For Each System In SystemSet
    text2.Text = text2.Text & System.Caption & ENTER
    text2.Text = text2.Text & System.Manufacturer & ENTER
    text2.Text = text2.Text & System.BuildType & ENTER
    text2.Text = text2.Text & "Version: " + System.Version & ENTER
    text2.Text = text2.Text & "Serial Number: " + System.SerialNumber & ENTER
Next
text2.Text = text2.Text & ENTER
End Sub
Public Sub GetProcessorData()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
'Processor Information
Set obj = GetObject("winmgmts:").InstancesOf("Win32_Processor")


            For Each obj2 In obj
            text3.Text = text3.Text & obj2.Caption & ENTER
            text3.Text = text3.Text & "Speed: " & obj2.currentclockspeed & " Mhz" & ENTER
Next
text3.Text = text3.Text & ENTER
End Sub
Public Sub GetMemoryData()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
'get memory
Set obj = GetObject("winmgmts:").InstancesOf("Win32_PhysicalMemory")
Dim i As String

            For Each obj2 In obj
            Text8.Text = obj2.capacity
            i = Text8.Text
            ii = i / 1024
            iii = ii / 1024
            text4.Text = text4.Text & iii & " MB" & " Chip" & ENTER
Next
text4.Text = text4.Text & ENTER
End Sub
Public Sub GetDriveData()
ENTER = Chr$(13) + Chr$(10)
'Drive Info
On Error GoTo driveerror
Set obj = GetObject("winmgmts:").InstancesOf("Win32_DiskDrive")

            For Each obj2 In obj
            Text8.Text = obj2.Size
            i = Text8.Text
            ii = i / 1024
            iii = ii / 1024
            iiii = iii / 1024
           text5.Text = text5.Text & obj2.Caption & " - " & Left$(iiii, 5) & " GB" & ENTER
          
Next
text5.Text = text5.Text & ENTER
  Exit Sub
driveerror:
  text5.Text = text5.Text & "Removable Drive"
text5.Text = text5.Text & ENTER
End Sub
Public Sub GetAdapterData()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
'Adapter information
    ' Create a sink to recieve the results of the enumeration
    Set sink = New SWbemSink
    
    ' Connect to root\cimv2.
    Set adapter = GetObject("winmgmts:")
' Perform the asynchronous enumeration of processes
adapter.InstancesOfAsync sink, "Win32_NetworkAdapter"

End Sub
Private Sub Command1_Click()
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
text2.Text = ""
text3.Text = ""
text4.Text = ""
text5.Text = ""
text6.Text = ""
Text7.Text = "0"
Text8.Text = ""

Call GetOSData
DoEvents
Call GetAdapterData
DoEvents
Call GetProcessorData
DoEvents
Call GetMemoryData
DoEvents
Call GetDriveData
DoEvents
text6.Text = text6.Text & ENTER
DoEvents
Call CompareData
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
DoEvents
If Check2.Value = 1 Then
Timer1.Enabled = True
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim intfile As Integer
Dim pass As String
Dim pass2 As String
Dim pass3 As String
Dim pass4 As String
intfile = FreeFile
  Open App.Path & "\Settings.ini" For Input As #intfile
  Input #intfile, pass
  Input #intfile, pass2
  Input #intfile, pass3
  Input #intfile, pass4
  TextServer.Text = pass
  TextPort.Text = pass2
  Check1.Value = pass3
  Check2.Value = pass4
  Close #intfile
DoEvents
DoEvents


Call LoadStoredData
DoEvents
DoEvents
If Check1.Value = 1 Then
Timer2.Enabled = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\Settings.ini" For Output Access Write As #filehandle%
    Print #filehandle%, TextServer.Text
    Print #filehandle%, TextPort.Text
    Print #filehandle%, Check1.Value
    Print #filehandle%, Check2.Value
    Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\Settings.ini" For Output Access Write As #filehandle%
    Print #filehandle%, TextServer.Text
    Print #filehandle%, TextPort.Text
    Print #filehandle%, Check1.Value
    Print #filehandle%, Check2.Value
    Close

End Sub

Private Sub menusettings_Click()
On Error Resume Next
frmpassword.Show
End Sub

Private Sub sink_OnObjectReady(ByVal objWbemObject As WbemScripting.ISWbemObject, ByVal objWbemAsyncContext As WbemScripting.ISWbemNamedValueSet)
On Error Resume Next
ENTER = Chr$(13) + Chr$(10)
Dim i As Integer
i = Text7.Text
Set adapter = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=" & i & "")
Description = adapter.Description

text6.Text = text6.Text & Description & ENTER

If IsNull(adapter.MACAddress) Then
    text6.Text = text6.Text & "No MAC Address" & ENTER
    text6.Text = text6.Text & "" & ENTER
Else
    text6.Text = text6.Text & "Mac: " & adapter.MACAddress & ENTER
    text6.Text = text6.Text & "" & ENTER
End If
 

Text7.Text = i + 1
End Sub
Private Sub menuexit_Click()
End
End Sub
Private Sub tcpClient_Close()
   Call TcpClient.Close
End Sub

Private Sub tcpClient_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
   Dim result As Integer
   result = MsgBox(Source & ": " & Description, _
      vbOKOnly, "TCP/IP Error")
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\Settings.ini" For Output Access Write As #filehandle%
    Print #filehandle%, TextServer.Text
    Print #filehandle%, TextPort.Text
    Print #filehandle%, Check1.Value
    Print #filehandle%, Check2.Value
    Close
End

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Command1_Click
Timer2.Enabled = False
End Sub
