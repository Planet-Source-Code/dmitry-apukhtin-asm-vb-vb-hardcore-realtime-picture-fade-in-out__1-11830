VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDIB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFOHEADER, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectA Lib "gdi32" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Public hdc As Long
Public lpRGB As Long

Private hBMO As Long
Private hDIB As Long
Private bmH As BITMAPINFOHEADER

Sub Clone(Optional pic As StdPicture = Nothing)
    Class_Terminate
    Dim BMP As BITMAP
    If Not pic Is Nothing Then
        GetObjectA pic.handle, Len(BMP), BMP
    End If
    With bmH
        .biSize = Len(bmH)
        .biWidth = 640 'BMP.bmWidth
        .biHeight = 480 'BMP.bmHeight
        .biPlanes = 1
        .biBitCount = 24
        .biSizeImage = ((.biWidth * 3 + 3) And &HFFFFFFFC) * .biHeight
    End With
    hdc = CreateCompatibleDC(GetDC(0))
    hDIB = CreateDIBSection(hdc, bmH, 0, lpRGB, 0, 0)
    If hDIB Then
        hBMO = SelectObject(hdc, hDIB)
        If Not pic Is Nothing Then
            Dim hDCt As Long, hBMOt As Long
            hDCt = CreateCompatibleDC(hdc)
            hBMOt = SelectObject(hDCt, pic.handle)
            StretchBlt hdc, 0, 0, 640, 480, hDCt, 0, 0, BMP.bmWidth, BMP.bmHeight, vbSrcCopy
            SelectObject hDCt, hBMOt
            DeleteDC hDCt
        End If
    Else
        Err.Raise -1, , "DIB failed to create"
    End If
End Sub

Sub Paint(ByVal toDC As Long, Optional ByVal X As Long, Optional ByVal Y As Long)
    BitBlt toDC, X, Y, bmH.biWidth, bmH.biHeight, hdc, 0, 0, vbSrcCopy
End Sub

Private Sub Class_Terminate()
    If hdc Then
        If hBMO Then DeleteObject SelectObject(hdc, hBMO)
        DeleteObject hdc
    End If
End Sub

Function RGBSize() As Long
    RGBSize = bmH.biWidth * bmH.biHeight
End Function
