Attribute VB_Name = "mdlScreenCapture"
Option Explicit
Option Private Module

Private Const RC_PALETTE As Long = &H100
Private Const RASTERCAPS As Long = 38
Private Const SIZEPALETTE As Long = 104

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type

Private Type LOGPALETTE
    palVersion As Integer
    palNumEntries As Integer
    'Enough for 256 colors
    palPalEntry(255) As PALETTEENTRY
End Type

Private Declare Function GetWindowDC Lib "USER32" ( _
        ByVal hWnd As Long) As Long

Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
        ByVal hDC As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib _
        "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long

Private Declare Function SelectObject Lib "GDI32" ( _
        ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function GetDeviceCaps Lib "GDI32" ( _
        ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long

Private Declare Function GetSystemPaletteEntries Lib _
        "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, _
        ByVal wNumEntries As Long, lpPaletteEntries _
        As PALETTEENTRY) As Long

Private Declare Function CreatePalette Lib "GDI32" ( _
        lpLogPalette As LOGPALETTE) As Long

Private Declare Function SelectPalette Lib "GDI32" ( _
        ByVal hDC As Long, ByVal hPalette As Long, _
        ByVal bForceBackground As Long) As Long

Private Declare Function RealizePalette Lib "GDI32" ( _
        ByVal hDC As Long) As Long

Private Declare Function BitBlt Lib "GDI32" ( _
        ByVal hDCDest As Long, ByVal XDest As Long, _
        ByVal YDest As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hDCSrc As Long, _
        ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
        As Long

Private Declare Function DeleteDC Lib "GDI32" ( _
        ByVal hDC As Long) As Long

Private Declare Function ReleaseDC Lib "USER32" ( _
        ByVal hWnd As Long, ByVal hDC As Long) As Long

Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Long 'As Picture

    On Error GoTo ErrorRoutineErr

    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim rc As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long

    Dim LogPal As LOGPALETTE

    'get device context for the window
    hDCSrc = GetWindowDC(hWndSrc)

    'Create a memory device context for the copy process
    hDCMemory = CreateCompatibleDC(hDCSrc)
    'Create a bitmap and place it in the memory DC
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    'get screen properties
    'Raster capabilities
    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
    'Palette support
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    'Size of palette
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)

    'If the screen has a palette, make a copy
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
    'Create a copy of the system palette
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        rc = GetSystemPaletteEntries(hDCSrc, 0, 256, _
                LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
    'Select the new palette into the memory
    'DC and realize it
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        rc = RealizePalette(hDCMemory)
    End If

    'Copy the image into the memory DC
    rc = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, _
            hDCSrc, LeftSrc, TopSrc, vbSrcCopy)

    'Remove the new copy of the  on-screen image
    'hBmp = SelectObject(hDCMemory, hBmpPrev)

    'If the screen has a palette get back the palette that was
    'selected in previously
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    'Release the device context resources back to the system
    rc = DeleteDC(hDCMemory)
    rc = ReleaseDC(hWndSrc, hDCSrc)

    'Call CreateBitmapPicture to create a picture
    'object from the bitmap and palette handles.
    'then return the resulting picture object.

    'MsgBox Len(hBmp) & " : " & Len(hPal)

    'Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)

    'CaptureWindow = PADL(Str(Len(hBmp)), 6) & Str(hBmp) & Str(hPal)

    CaptureWindow = hBmp
    
ErrorRoutineResume:
    Exit Function
ErrorRoutineErr:
    MsgBox "CaptureWindow" & Err & error
    Resume Next
End Function


