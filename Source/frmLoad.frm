VERSION 5.00
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "XLISTEN.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VRS"
   ClientHeight    =   1215
   ClientLeft      =   9825
   ClientTop       =   7230
   ClientWidth     =   2190
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   2190
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Enable/Disable the voice command"
      Top             =   720
      Width           =   1935
   End
   Begin VB.ListBox List3 
      Height          =   1425
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Text1"
      ToolTipText     =   "The command recognised"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Add a voice command"
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Remove and see the recognised commands"
      Top             =   480
      Width           =   975
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR DirectSR1 
      Height          =   495
      Left            =   720
      OleObjectBlob   =   "frmLoad.frx":044A
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Sub DirectSR1_PhraseFinish(ByVal flags As Long, ByVal beginhi As Long, ByVal beginlo As Long, ByVal endhi As Long, ByVal endlo As Long, ByVal Phrase As String, ByVal parsed As String, ByVal results As Long)
    Dim noth
    List1.ListIndex = -1
        
    For i = 0 To List1.ListCount
        If Phrase = "" Then
            List1.ListIndex = -1
            List2.ListIndex = -1
            List3.ListIndex = -1
            Text1 = "Not Recognised"
            Exit Sub
        End If
        If Phrase = List1.List(i) Then
            List1.ListIndex = i
            List2.ListIndex = i
            List3.ListIndex = i
            Text1 = List3.List(i)
            noth = Shell(List2.List(i), vbNormalNoFocus)
        End If
    Next i
End Sub

Private Sub Form_Load()
    Dim junk, windir$
                   
    Command1.Caption = "Words List"
    Command2.Caption = "Add Word"
    Command3.Caption = "Disable"
    
    Text1.Text = ""
        
    windir = Space(144)
    junk = getwindir(windir, 144)
    windir = Trim(windir)
    i = InStr(windir$, vbNullChar)
    windir$ = Mid$(windir$, 1, i - 1)
    
    words = windir$ & "\words.txt"
    dirs = windir$ & "\dirs.txt"
    descrip = windir$ & "\desc.txt"
    
    test = Dir(words)
    If test = "" Then
        Open words For Output As #1
        Close #1
    End If
    
    test = Dir(dirs)
    If test = "" Then
        Open dirs For Output As #1
        Close #1
    End If
    
    test = Dir(descrip)
    If test = "" Then
        Open descrip For Output As #1
        Close #1
    End If
    
    Call loadfiles
    Text1 = "Ready"
        
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call savefiles
    Set Form1 = Nothing
    Set Form2 = Nothing
    Set Form3 = Nothing
    End
End Sub

Private Sub command1_click()
    Form2.Show
End Sub

Private Sub Command2_Click()
    Form3.Show
End Sub

Private Sub Command3_Click()
    Select Case Command3.Caption
    Case Is = "Disable"
        DirectSR1.Deactivate
        Command3.Caption = "Enable"
        Text1 = "Disabled"
    Case Is = "Enable"
        DirectSR1.Activate
        Command3.Caption = "Disable"
        Text1 = "Ready"
    End Select
End Sub
