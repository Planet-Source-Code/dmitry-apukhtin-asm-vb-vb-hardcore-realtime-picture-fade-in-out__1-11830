VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Speed it UP!"
   ClientHeight    =   7620
   ClientLeft      =   1380
   ClientTop       =   510
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Faster"
      Height          =   495
      Left            =   8175
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox tCPU 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "?"
      Top             =   7035
      Width           =   2820
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CPU ID ->"
      Height          =   315
      Left            =   180
      TabIndex        =   2
      Top             =   7020
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modal Form"
      Height          =   495
      Left            =   8100
      TabIndex        =   1
      Top             =   6915
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Text            =   "Calculating FPS..."
      Top             =   0
      Width           =   2820
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   oh, boy... I'm so tired to write any comments...
'   anyway, cool stuff, eh? :-) I'd say it's _too_ cool ;-)
'   if you could understand how all this crap works -
'   count yourself as a real profy ;-)
'   otherwise, hope you'll learn some tricks from this source.
'
'   regards,
'   /Damian
'   dmitrya@thewercs.com

Dim DIB As CDIB, DIBsrc As CDIB

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Asm As Asm
Private lFader As Long, lDelta As Long
Private fCount As Long, tm As Long
Private rcSS As RECT, SS As String
Private curFile As String

Private Sub Check1_Click()
    Check1.Refresh
    lDelta = lDelta Xor 8
End Sub

Private Sub Command1_Click()
    frmAsmClass.Show vbModal
End Sub

Private Sub Command2_Click()
    Dim ts As String: ts = Space$(12)
    Asm.CpuID ts
    tCPU = ts
    tCPU.Refresh
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Open App.Path & "\Asm.cls" For Input As #1
    SS = Input$(LOF(1), 1)
    Close #1
    SetRect rcSS, 10, 800, 610, 430
    On Error GoTo E
    
    Set DIB = New CDIB
    DIB.Clone
    Set DIBsrc = New CDIB
    Set Asm = New Asm
    lFader = 4
    lDelta = -4
    Show
    Refresh
Q:
    Exit Sub
E:
    MsgBox Error, vbCritical
    Unload Me
    Debug.Assert 0
    Resume Q
    Resume
End Sub

Private Sub Form_Paint()
    lFader = lFader + lDelta
    If lFader > 255& - lDelta Then
        lDelta = -lDelta
    ElseIf lFader < -lDelta Then
        If curFile = "" Then curFile = Dir$(App.Path & "\*.jpg")
        DIBsrc.Clone LoadPicture(App.Path & "\" & curFile)
        curFile = Dir$
        SetBkMode DIB.hdc, 1
        SetTextColor DIB.hdc, vbWhite
        
        lDelta = -lDelta
        tm = timeGetTime + 1000
        fCount = 0
    End If
    Asm.XFade DIBsrc.lpRGB, DIB.lpRGB, DIB.RGBSize, lFader
    rcSS.Top = rcSS.Top - 1
    DrawText DIB.hdc, SS, -1, rcSS, &H900
    DIB.Paint hdc, 0, 25
    Dim rc As RECT: rc.Right = 1: rc.Bottom = 1
    fCount = fCount + 1
    If timeGetTime > tm Then
        DoEvents
        Text1 = fCount & " FPS": Text1.Refresh
        tm = timeGetTime + 1000
        fCount = 0
        If rcSS.Top < -2600 Then rcSS.Top = 430
    End If
    InvalidateRect hWnd, rc, False
End Sub

Private Sub Form_Resize()
    Text1.Move 0, 0, ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Asm = Nothing
End Sub

