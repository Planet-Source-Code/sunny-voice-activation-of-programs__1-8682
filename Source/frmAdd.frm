VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Word"
   ClientHeight    =   1860
   ClientLeft      =   4680
   ClientTop       =   4275
   ClientWidth     =   4020
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4020
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2040
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text3"
      ToolTipText     =   "Locate the executable"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text2"
      ToolTipText     =   "A short description here"
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   $"frmAdd.frx":044A
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
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

Private Sub Form_Load()

    Label1.Caption = "Sound:"
    Label2.Caption = "Description:"
    Label3.Caption = "File to execute:"
    
    Command1.Caption = "Save"
    Command2.Caption = "Close"
    Command3.Caption = "Browse"
    
    Text1 = ""
    Text2 = ""
    Text3 = ""
    
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub command1_click()
    Dim count, pos, i, length As Integer
    Dim line, test As Variant
        
    Form1.List1.AddItem (Text1.Text)
    Form1.List2.AddItem (Text3.Text)
    Form1.List3.AddItem (Text2.Text)
    
    Form2.List1.AddItem (Text1.Text)
    Form2.List2.AddItem (Text2.Text)
    
'    count = 0: pos = 0
'    Do
'        pos = InStr(pos + 1, Text3.Text, "\")    'find how many "\"s there are
'        If pos = 0 Then Exit Do            'in the full dir
'        count = count + 1
'    Loop Until pos = 0
'
'    For i = 0 To count - 1                  'using the last result, gets
'        pos = InStr(pos + 1, Text3.Text, "\") + 1 'position of the last "\"
'    Next i                                  'and add 1 so it is pos after "\"'
'
'    length = Len(Text3.Text) - pos + 1            'length of filetitle
'    line = Mid(Text3.Text, pos, length)           'extract filetitle
    
'    Form2.List3.AddItem (line)
    
    Call savefiles
    Form1.List1.Clear
    Form1.List2.Clear
    Form1.List3.Clear
    Form2.List1.Clear
    Form2.List2.Clear
    Form2.List3.Clear
    Call loadfiles
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    CommonDialog1.Filter = "Executable Files (*.exe)|*.EXE"
    CommonDialog1.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    
    Text3 = CommonDialog1.FileName
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form3 = Nothing
End Sub
