Attribute VB_Name = "mdlScreenCapture"
Option Explicit
Option Base 0

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

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type


Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PicBmp
   Size As Long
   bitMapType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Declare Function BitBlt Lib "GDI32" ( _
        ByVal hDCDest As Long, ByVal XDest As Long, _
        ByVal YDest As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long, ByVal hDCSrc As Long, _
        ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) _
        As Long
Private Declare Function CreateCompatibleBitmap Lib _
        "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, _
        ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" ( _
        ByVal hDC As Long) As Long
Private Declare Function CreatePalette Lib "GDI32" ( _
        lpLogPalette As LOGPALETTE) As Long
Private Declare Function DeleteDC Lib "GDI32" ( _
        ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" ( _
        ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib _
        "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, _
        ByVal wNumEntries As Long, lpPaletteEntries _
        As PALETTEENTRY) As Long
Private Declare Function GetWindowDC Lib "USER32" ( _
        ByVal hWnd As Long) As Long
Private Declare Function OleCreatePictureIndirect _
        Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
        ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function RealizePalette Lib "GDI32" ( _
        ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "USER32" ( _
        ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "GDI32" ( _
        ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SelectPalette Lib "GDI32" ( _
        ByVal hDC As Long, ByVal hPalette As Long, _
        ByVal bForceBackground As Long) As Long

'Public Function CaptureForm(frmSrc As Form) As Picture
'    On Error GoTo ErrorRoutineErr
'
'    'Call CaptureWindow to capture the entire form
'    'given it's window
'    'handle and then return the resulting Picture object
'    Set CaptureForm = CaptureWindow(frmSrc.hWnd, 0, 0, _
'            frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), _
'            frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
'
'ErrorRoutineResume:
'    Exit Function
'ErrorRoutineErr:
'    MsgBox "Project1.Module1.CaptureForm" & Err & Error
'    Resume Next
'End Function

Public Function CreateBitmapPicture(ByVal hBmp As Long, _
    ByVal hPal As Long) As Picture
    
    On Error GoTo ErrorRoutineErr
    
    Dim r As Long
    Dim Pic As PicBmp
    'IPicture requires a reference to "Standard OLE Types"
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    'Fill in with IDispatch Interface ID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    'Fill Pic with necessary parts
    With Pic
    'Length of structure
        .Size = Len(Pic)
    'Type of Picture (bitmap)
        .bitMapType = vbPicTypeBitmap
    'Handle to bitmap
        .hBmp = hBmp
    'Handle to palette (may be null)
        .hPal = hPal
    End With

    'Create Picture object
    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)

    'Return the new Picture object
    Set CreateBitmapPicture = IPic

ErrorRoutineResume:
    Exit Function
ErrorRoutineErr:
    MsgBox "Project1.Module1.CreateBitmapPicture" & Err & Error
    Resume Next
End Function

Public Function CaptureWindow(ByVal hWndSrc As Long, _
    ByVal LeftSrc As Long, _
    ByVal TopSrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long) As Picture

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
    
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)

ErrorRoutineResume:
    Exit Function
ErrorRoutineErr:
    MsgBox "CaptureWindow" & Err & Error
    Resume Next
End Function

Public Sub PrintPicture(Prn As Printer, Pic As Picture)
    On Error GoTo ErrorRoutineErr
'
' Prints out the selected picture to the printer
'
    Prn.PaintPicture Pic, 0, 0

ErrorRoutineResume:
    Exit Sub
ErrorRoutineErr:
    MsgBox "PrintPicture" & Err & Error
    Resume Next
End Sub

Public Function PADL(ByVal text As String, length As Integer, Optional Symbol As String = " ") As String
    Dim tmpLen As Integer

    PADL = Trim(text)

    Symbol = Mid(Symbol, 1, 1)

    tmpLen = Len(text)

    Do While (tmpLen < length)
        PADL = PADL & Symbol
        tmpLen = tmpLen + 1
    Loop

End Function

Public Function PADR(ByVal text As String, length As Integer, Optional Symbol As String = " ") As String
    Dim tmpLen As Integer

    PADR = Trim(text)

    Symbol = Mid(Symbol, 1, 1)

    tmpLen = Len(text)

    Do While (tmpLen < length)
        PADR = Symbol & PADR
        tmpLen = tmpLen + 1
    Loop

End Function

