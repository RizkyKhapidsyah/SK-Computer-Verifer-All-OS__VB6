VERSION 5.00
Begin VB.Form frmsettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   Icon            =   "frmsettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   5205
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox textpassword 
      Height          =   285
      Left            =   3240
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   6000
   End
   Begin VB.TextBox TextPort 
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Text            =   "1979"
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Default"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox TextServer 
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Text            =   "Localhost"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Current information to store"
      Height          =   3975
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   4935
      Begin VB.TextBox text2 
         DataField       =   "Os"
         Height          =   765
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   240
         Width           =   3375
      End
      Begin VB.TextBox text3 
         DataField       =   "Processor Info"
         Height          =   525
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox text4 
         DataField       =   "Memory"
         Height          =   525
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox text5 
         DataField       =   "Drives"
         Height          =   765
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox text6 
         DataField       =   "Adapter Info"
         Height          =   765
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3120
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Os:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Processor Info:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Memory:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Drives:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Adapter Info:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   3120
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Password protect ""Settings"""
      Height          =   1575
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "Set"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox textconfirm 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox textpass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Confirm Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Store computer information"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   6000
      Width           =   2415
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   1200
      TabIndex        =   19
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2520
      TabIndex        =   20
      Text            =   "0"
      Top             =   5760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Port"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Server IP or Name"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmsettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents sink As SWbemSink
Attribute sink.VB_VarHelpID = -1
Function Encrypt(What As String) As String
    Dim Before$, After$, EpN%, Dracula%, Aeneima$, DeMoNs$
    Before$ = " ?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZø°≤≥¿¡¬√ƒ≈“”‘’÷Ÿ€‹‡·‚„‰Âÿ∂ß⁄•"
    After$ = " ø°@#$%^&*()_+|01≤≥456789¿b¡d¬√ghƒjklm≈“”q‘’÷Ÿvw€‹z.-~,A‡·‚„FGH‰JKÂMNÿ∂QRßT⁄VWX•Z?!23acefinoprstuxyBCDEILOPSUY"
    For EpN% = 1 To Len(What)
        Dracula% = InStr(Before$, Mid(What, EpN%, 1))
        If Not Dracula% = 0 Then
            Aeneima$ = Mid(After$, Dracula%, 1)
            DeMoNs$ = DeMoNs$ + Aeneima$
        End If
    Next
    Encrypt = DeMoNs$
End Function
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

  Exit Sub
driveerror:
  text5.Text = text5.Text & "Removable Drive"

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
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\OS.dat" For Output Access Write As #filehandle%
    Print #filehandle%, text2.Text
    Close

    Open App.Path & "\Processor.dat" For Output Access Write As #filehandle%
    Print #filehandle%, text3.Text
    Close

    Open App.Path & "\Memory.dat" For Output Access Write As #filehandle%
    Print #filehandle%, text4.Text
    Close

    Open App.Path & "\Drives.dat" For Output Access Write As #filehandle%
    Print #filehandle%, text5.Text
    Close

    Open App.Path & "\Adapter.dat" For Output Access Write As #filehandle%
    Print #filehandle%, text6.Text
    Close
    
Call frmmain.LoadStoredData
End Sub

Private Sub Command2_Click()
On Error Resume Next
If textpass.Text = textconfirm.Text Then
textpassword.Text = Encrypt(textconfirm.Text)
DoEvents
DoEvents
DoEvents
Dim filehandle%
    filehandle = FreeFile
    Open App.Path & "\Data.dat" For Output Access Write As #filehandle%
    Print #filehandle%, textpassword.Text
    Close
Else
MsgBox "Sorry you mistyped your password please try again."
End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
TextServer.Text = "Localhost"
TextPort.Text = "1979"
End Sub

Private Sub Form_Load()
On Error Resume Next
frmsettings.TextServer.Text = frmmain.TextServer.Text
frmsettings.TextPort.Text = frmmain.TextPort.Text
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

Private Sub Command5_Click()
frmmain.TextServer.Text = frmsettings.TextServer.Text
frmmain.TextPort.Text = frmsettings.TextPort.Text
Unload Me
End Sub

Private Sub textconfirm_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
 Call Command2_Click
 DoEvents
 End If

End Sub

Private Sub textpass_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Then
 textconfirm.SetFocus
 DoEvents
 End If

End Sub

Private Sub Timer1_Timer()
text2.Text = ""
text3.Text = ""
text4.Text = ""
text5.Text = ""
text6.Text = ""
Text7.Text = "0"
Text8.Text = ""

Call GetOSData
DoEvents

Call GetProcessorData
DoEvents

Call GetMemoryData
DoEvents

Call GetDriveData
DoEvents

Call GetAdapterData
DoEvents
Timer1.Enabled = False
End Sub
