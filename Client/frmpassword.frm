VERSION 5.00
Begin VB.Form frmpassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Password"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2310
   Icon            =   "frmpassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   2310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "If you have not set a password yet then there is no password. Just hit Ok."
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Please enter password:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function DeEncrypt(What As String) As String
    Dim Before$, After$, EpN%, Dracula%, Aeneima$, DeMoNs$
    Before$ = " ø°@#$%^&*()_+|01≤≥456789¿b¡d¬√ghƒjklm≈“”q‘’÷Ÿvw€‹z.-~,A‡·‚„FGH‰JKÂMNÿ∂QRßT⁄VWX•Z?!23acefinoprstuxyBCDEILOPSUY"
    After$ = " ?!@#$%^&*()_+|0123456789abcdefghijklmnopqrstuvwxyz.,-~ABCDEFGHIJKLMNOPQRSTUVWXYZø°≤≥¿¡¬√ƒ≈“”‘’÷Ÿ€‹‡·‚„‰Âÿ∂ß⁄•"
    For EpN% = 1 To Len(What)
        Dracula% = InStr(Before$, Mid(What, EpN%, 1))
        If Not Dracula% = 0 Then
            Aeneima$ = Mid(After$, Dracula%, 1)
            DeMoNs$ = DeMoNs$ + Aeneima$
        End If
    Next
    DeEncrypt = DeMoNs$
End Function
Private Sub Command1_Click()
If Text1.Text = text2.Text Then
frmsettings.Show
Unload Me
Else
MsgBox "Wrong password. Please try again"
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim intfile As Integer
Dim pass As String
intfile = FreeFile
  Open App.Path & "\Data.dat" For Input As #intfile
  Input #intfile, pass
  text3.Text = pass
  Close #intfile
DoEvents
DoEvents
text2.Text = DeEncrypt(text3.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call Command1_Click
 DoEvents
 End If

End Sub
