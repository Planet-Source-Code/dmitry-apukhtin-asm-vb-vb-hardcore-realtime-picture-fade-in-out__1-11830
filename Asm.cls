VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Asm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'   Asmippets by Damian (Assembler Snippets)
'   the way of calling procedures in machine codes from VB
'
'   well, it's so hard to comment this stuff... if you're really eager to,
'   you'd carefully step thru each single line of it and read tons
'   of MSDN articles. if not, heh, it wasn't worth to comment it.
'
'   general idea - replace VTable entries of class factory interface
'   by links to pre-allocated memory blocks with our code.
'   not a fastest way but works in both IDE and EXE, p-code and native.
'
'   fastest way is replacing proc bodies in global modules - eliminates
'   overhead of class methods calling and COM parameters passing. but
'   no way to make it work in IDE except of using stuff like
'   CallWindowProc - but this gets us back to call overhead and screws
'   parameters passing.
'
Private Type TPROC
    hMem As Long
    vtPtr As Long
End Type
Private aProc() As TPROC

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Sub XFade(ByVal lpSrc As Long, ByVal lpDest As Long, ByVal Size As Long, ByVal Amount As Long)
    ' loops for each byte:
    '   <lpDest>=LOBYTE(<lpSrc>*Amount)
    I_AM "5756 8B7424108B7C24148B4C2418C1E9028B54241CACF6E28AC4AAACF6E28AC4AAACF6E28AC4AAACF6E28AC4AAAC" & _
         "F6E28AC4AAACF6E28AC4AAACF6E28AC4AAACF6E28AC4AAACF6E28AC4AAACF6E28AC4AAACF6E28AC4AAACF6E28AC4AAE2B6" & _
         "5E5F 33C0C21400"
End Sub

Sub RMChar(ByRef str As String, Optional ByVal CharAscii As Long = 32)
    ' removes each occurence of CharAscii in str
    I_AM "57568B7C24108B0785C074508BF88B4FFC668B4424146689040FD1E941F266AF752A4F4F418BF7F366AF74204F4F418BD18BD" & _
         "FF266AF4F4F412BCAF7D92BD18BFE8BF3F366A58BCA87F7EBDBC606008BC68B7C24108B3F2BC78947FC5E5F 33C0C20C00"
End Sub

Function Peek(ByVal adr As Long) As Long
    ' peeks a long from absolute adr
    I_AM "8B4C2408 8B54240C 8B01 8902 33C0 C20C00"
End Function

Sub Movs(ByVal lpDest As Long, ByVal lpSource As Long, ByVal ByteCount As Long)
    ' executes REP MOVSD (and REP MOVSB for last 3 bytes if ByteCount not even)
    I_AM "8BDF8BD6 8B7C2408 8B74240C 8B4C2410 8BC1 C1E902 F3A5 8BC8 83E103 F3A4 8BF28BFB 33C0 C21000"
End Sub

Sub CpuID(ByVal strBuf As String)
    ' stores CPUID in strBuf (with converting ASCII->UNICODE)
    ' strBuf must be pre-allocated to at least of 12 characters
    I_AM "578B7C240C33C00FA28AC366AB8AC766ABC1EB108AC366AB8AC766AB8BDA8AC366AB8AC766ABC1EB108AC366AB8AC766AB8BD98AC366AB8AC766ABC1EB108AC366AB8AC766AB5F33C0C20800"
End Sub

Private Sub Class_Initialize()
    Dim a As Long, b As Long
    Dim ts As String * 12
    
    XFade VarPtr(a), VarPtr(b), 4, 0
    RMChar vbNullString
    Peek VarPtr(a)
    Movs VarPtr(b), VarPtr(a), 4
    CpuID ts
End Sub

'
'
'
Private Sub I_AM(s As String)
    s = Replace$(s, " ", "")
    Dim i As Long, aSize As Long, aB() As Byte
    aSize = Len(s) \ 2

    ReDim Preserve aB(1 To aSize)
    For i = 1 To aSize
        aB(i) = Val("&H" & Mid$(s, i * 2 - 1, 2))
    Next
    
    Static cp As Long
    ReDim Preserve aProc(cp)
    Dim hMem As Long, lPtr As Long
    hMem = GlobalAlloc(0, aSize)
    lPtr = GlobalLock(hMem)
    CopyMemory ByVal lPtr, aB(1), aSize
    GlobalUnlock hMem
    
    aProc(cp).hMem = hMem
    aProc(cp).vtPtr = VTable(cp)
    VTable(cp) = lPtr
    cp = cp + 1
End Sub

Private Sub Class_Terminate()
    On Error GoTo E
    Dim i As Long
    For i = 0 To UBound(aProc)
        VTable(i) = aProc(i).vtPtr
        GlobalFree aProc(i).hMem
    Next
E:
End Sub

Private Property Get VTable(ByVal Index As Long) As Long
    Dim p As Long
    Index = &H1C + Index * 4
    CopyMemory p, ByVal ObjPtr(Me), 4
    CopyMemory VTable, ByVal p + Index, 4
End Property

Private Property Let VTable(ByVal Index As Long, ByVal ProcPtr As Long)
    Dim p As Long
    Index = &H1C + Index * 4
    CopyMemory p, ByVal ObjPtr(Me), 4
    CopyMemory ByVal p + Index, ProcPtr, 4
End Property

