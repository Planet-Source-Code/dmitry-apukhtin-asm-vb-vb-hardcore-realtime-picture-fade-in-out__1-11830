VERSION 5.00
Begin VB.Form frmAsmClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "asm class"
   ClientHeight    =   2295
   ClientLeft      =   3720
   ClientTop       =   2580
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4635
   Begin VB.TextBox tRmChar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   12
      Top             =   840
      Width           =   705
   End
   Begin VB.TextBox tReplace 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1515
      TabIndex        =   11
      Top             =   840
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RMChar ->"
      Height          =   525
      Left            =   105
      TabIndex        =   9
      Top             =   1695
      Width           =   1320
   End
   Begin VB.TextBox tCopyMem 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1515
      TabIndex        =   0
      Top             =   135
      Width           =   750
   End
   Begin VB.TextBox tCopyMemBig 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1515
      TabIndex        =   1
      Top             =   480
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Time It!"
      Default         =   -1  'True
      Height          =   360
      Left            =   75
      TabIndex        =   4
      Top             =   1185
      Width           =   4500
   End
   Begin VB.TextBox tPeek 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   2
      Top             =   135
      Width           =   705
   End
   Begin VB.TextBox tMovsD 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   3
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblRMS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1500
      TabIndex        =   15
      Top             =   1695
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "RmChar"
      Height          =   195
      Index           =   5
      Left            =   2760
      TabIndex        =   14
      Top             =   870
      Width           =   570
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "native Replace"
      Height          =   195
      Index           =   2
      Left            =   345
      TabIndex        =   13
      Top             =   870
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   0
      X2              =   6000
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   6000
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Label lblRMChar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1500
      TabIndex        =   10
      Top             =   1950
      Width           =   3060
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "4-byte CopyMem:"
      Height          =   195
      Index           =   3
      Left            =   195
      TabIndex        =   8
      Top             =   180
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "32-byte CopyMem:"
      Height          =   195
      Index           =   4
      Left            =   105
      TabIndex        =   7
      Top             =   510
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "cpyMem"
      Height          =   195
      Index           =   1
      Left            =   2730
      TabIndex        =   6
      Top             =   510
      Width           =   600
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Peek:"
      Height          =   195
      Index           =   0
      Left            =   2910
      TabIndex        =   5
      Top             =   180
      Width           =   420
   End
End
Attribute VB_Name = "frmAsmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LOOPS = 5000000
Const strSpaces = " 0      1 2  3   4    5     6 789 -10- 98765 4321 !"

Private tm As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim bigArray() As Byte

Private Asm As Asm

Private Sub Command1_Click()
Form1.AutoRedraw = True
    On Error GoTo E
    Command1.Enabled = False
    
    Screen.MousePointer = vbHourglass
    
    Dim L1 As Long, L2 As Long, i As Long
    L2 = VarPtr(L2)
    
    tmStart tCopyMem
    For i = 1 To LOOPS
        CopyMemory L1, L2, 4
    Next
    tmDone tCopyMem

    tmStart tPeek
    For i = 1 To LOOPS
        L1 = Asm.Peek(L2)
    Next
    tmDone tPeek

    Dim hMem As Long, memSize As Long
    Dim lPtr As Long, aPtr As Long
    memSize = 32
    ReDim bigArray(1 To memSize)
    hMem = GlobalAlloc(0, UBound(bigArray))
    lPtr = GlobalLock(hMem)
    aPtr = VarPtr(bigArray(1))
    
    tmStart tCopyMemBig
    For i = 1 To LOOPS
        CopyMemory ByVal aPtr, ByVal lPtr, memSize
    Next
    tmDone tCopyMemBig

    tmStart tMovsD
    L2 = VarPtr(L2)
    For i = 1 To LOOPS
        Asm.Movs aPtr, lPtr, memSize
    Next
    tmDone tMovsD

    GlobalUnlock hMem
    GlobalFree hMem

    Dim ts As String
    tmStart tReplace
    For i = 1 To 30000
        ts = Replace$(strSpaces, " ", "")
    Next
    tmDone tReplace

    tmStart tRmChar
    L2 = VarPtr(L2)
    For i = 1 To 30000
        ts = strSpaces
        Asm.RMChar ts
    Next
    tmDone tRmChar

Q:
    Command1.Enabled = True
    Screen.MousePointer = 0
Form1.AutoRedraw = False
Form1.Refresh
    Exit Sub
E:
    MsgBox "Check what you entered", vbCritical, "Hey"
    Resume Q
    Resume
End Sub

Sub tmStart(T As TextBox)
    T = "Working"
    DoEvents
    Dim i As Long
    i = timeGetTime
    Do
        tm = timeGetTime
    Loop While i = tm
End Sub

Sub tmDone(T As TextBox)
    T = timeGetTime - tm
    DoEvents
End Sub

Private Sub Command2_Click()
    Dim ts As String
    ts = lblRMS
    Asm.RMChar ts
    lblRMChar = ts
End Sub

Private Sub Form_Initialize()
    Set Asm = New Asm
End Sub

Private Sub Form_Load()
    lblRMS_Click
End Sub

Private Sub Form_Terminate()
    Set Asm = Nothing
End Sub

Private Sub Label1_Click(Index As Integer)
    Dim ts As String
    ts = strSpaces
    Asm.RMChar ts
    MsgBox strSpaces & vbCr & ts, vbInformation, "RMChar test"
End Sub

Private Sub lblRMS_Click()
    Dim i As Long
    lblRMS = ""
    For i = 1 To 10
        lblRMS = lblRMS & String$(3 * Rnd + 1, Chr(65 + 26 * Rnd))
        If Rnd > 0.7 Then lblRMS = lblRMS & " "
        If Rnd > 0.7 Then lblRMS = lblRMS & " "
        If Rnd > 0.7 Then lblRMS = lblRMS & " "
    Next
End Sub
