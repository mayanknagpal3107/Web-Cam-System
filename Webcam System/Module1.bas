Attribute VB_Name = "Module1"
Option Explicit

Private Const API_FALSE As Long = 0
Private Const API_TRUE As Long = 1
Private Const API_NULL As Long = 0

Private Const CBM_INIT As Long = &H4&

Private Const DIB_RGB_COLORS As Long = 0
Private Const DIB_PAL_COLORS As Long = 1

Private Enum BiCompressionValues
    BI_RGB = 0
    BI_BITFIELDS = 3
    BI_FOURCC_YUY2 = &H32595559
    BI_FOURCC_UYVY = &H59565955
End Enum

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As BiCompressionValues
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As Long 'RGBQUAD
End Type

Private Type PictDescBmp
    cbSizeOfStruct As Long
    picType As PictureTypeConstants
    hBitmap As Long
    hPal As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" ( _
    ByVal hDC As Long, _
    ByVal pbmi As Long, _
    ByVal iUsage As Long, _
    ByRef ppvBits As Long, _
    ByVal hSection As Long, _
    ByVal dwOffset As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GdiFlush Lib "gdi32" () As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, _
    ByRef pSource As Any, _
    ByVal Length As Long)

Private Declare Function OleCreatePictureIndirect Lib "olepro32" ( _
    ByRef pPicDesc As PictDescBmp, _
    ByRef RefIID As GUID, _
    ByVal fOwn As Long, _
    ByRef IPic As IPicture) As Long

Public Function LongDIB2HBitmap(ByRef LongDIB() As Long) As Long
    'Returns non-0 on success.
    Dim hMemDC As Long
    Dim bmiHeader As BITMAPINFOHEADER 'Data copied into here for access.
    Dim ColorOffset As Long
    Dim BitsOffset As Long
    Dim Usage As Long
    Dim pBits As Long
    Dim SizeImage As Long
    
    hMemDC = CreateCompatibleDC(API_NULL)
    If hMemDC Then
        MoveMemory bmiHeader, LongDIB(0), Len(bmiHeader)
        With bmiHeader
            ColorOffset = .biSize \ 4
            'We have a "packed DIB" so biClrUsed is either 0 or
            'the actual color table size!
            BitsOffset = ColorOffset + .biClrUsed
            If .biClrUsed Then Usage = DIB_PAL_COLORS
            Select Case .biCompression
                Case BI_RGB, BI_BITFIELDS
                    'Pad scan line to full width, multiply by height.
                    SizeImage = ((((.biWidth * .biBitCount) + &H1F) And Not &H1F&) \ &H8) _
                              * Abs(.biHeight)
                Case Else
                    SizeImage = .biSizeImage
            End Select
            LongDIB2HBitmap = CreateDIBSection(hMemDC, _
                                               VarPtr(LongDIB(0)), _
                                               Usage, _
                                               pBits, _
                                               API_NULL, _
                                               0)
            If LongDIB2HBitmap <> 0 Then
                GdiFlush
                MoveMemory ByVal pBits, LongDIB(BitsOffset), SizeImage
            End If
        End With
        DeleteDC hMemDC
    End If
End Function

Public Function HBitmap2Picture(ByVal hBitmap As Long, ByVal hPal As Long) As Picture
    'Returns Nothing on failure.
    Dim Result As Long
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
    Dim BmpDesc As PictDescBmp
    
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    With BmpDesc
        .cbSizeOfStruct = Len(BmpDesc)
        .picType = vbPicTypeBitmap
        .hBitmap = hBitmap
        .hPal = hPal
    End With
    
    Result = OleCreatePictureIndirect(BmpDesc, IID_IDispatch, API_TRUE, IPic)
    If Result = 0 Then
        Set HBitmap2Picture = IPic
    End If
End Function


