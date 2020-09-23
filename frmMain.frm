VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bin In Exe"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGen 
      Caption         =   "Generate"
      Height          =   315
      Left            =   960
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cdmSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   855
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Any File (*.*)|*.*|"
      Flags           =   7
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Text            =   "File1"
      Top             =   840
      Width           =   3375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "File Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   880
      Width           =   750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cdmSave_Click()
Dim s$, l&
 l& = FreeFile()
 Open txtPath(0).Text For Binary Access Read As l&
  s$ = Input(LOF(l&), #l&)
 Close l&

 l& = FreeFile()
 Open txtPath(1).Text For Binary Access Read As l&
  s$ = Input(LOF(l&), #l&) & "&" & txtName.Text & "&" & s$ & "&/" & txtName.Text & "&"
 Close l&

 l& = FreeFile()
 Open txtPath(1).Text For Binary Access Write As l&
  Put #l&, 1, s$
 Close l&
End Sub

Private Sub cmdBrowse_Click(Index As Integer)
On Error GoTo 1
 If Index% = 0 Then cd.Filter = "Any File (*.*)|*.*|" Else _
    cd.Filter = "Executable (*.exe)|*.exe|Any File (*.*)|*.*|"
 
 cd.ShowOpen
 If cd.FileName <> "" Then txtPath(Index%).Text = cd.FileName Else Exit Sub
 If Index% = 0 Then txtName.Text = Mid$(cd.FileName, InStrRev(cd.FileName, "\") + 1)
1
End Sub

Private Sub cmdGen_Click()

If MsgBox("Ok to copy to Clipboard?", vbQuestion + vbYesNo, "Generated") = vbNo Then Exit Sub

Dim s$
s$ = s$ & "Dim s$, l&" & vbCrLf
s$ = s$ & "Dim t&, y&" & vbCrLf
s$ = s$ & " l& = FreeFile()" & vbCrLf
s$ = s$ & " Open App.Path & IIf(Right$(App.Path, 1) <> ""\"", ""\"", """") & App.EXEName & "".exe"" For Binary Access Read As l&" & vbCrLf
s$ = s$ & "  s$ = Input(LOF(l&), #l&)" & vbCrLf
s$ = s$ & " Close l&" & vbCrLf
s$ = s$ & vbCrLf
s$ = s$ & " t& = InStr(s$, ""&" & txtName.Text & "&"")" & vbCrLf
s$ = s$ & " y& = InStr(t& + 1, s$, ""&/" & txtName.Text & "&"")" & vbCrLf
s$ = s$ & " If t& = 0 Or y& = 0 Then Exit Sub" & vbCrLf
s$ = s$ & " s$ = Mid$(s$, t& + Len(""&" & txtName.Text & "&""), y& - t& - Len(""&/" & txtName.Text & "&""))" & vbCrLf

Clipboard.Clear
Clipboard.SetText s$
End Sub
