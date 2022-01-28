Attribute VB_Name = "mdlBMPCreator"
Option Explicit
Option Private Module

Private Type PicBmp
   Size As Long
   bitMapType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Declare Function OleCreatePictureIndirect _
        Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
        ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Public Function CreateBitmapPicture(ByVal hBmp As Long) As Picture

    On Error GoTo ErrorRoutineErr

    Dim hPal As Long
    
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

    hPal = 0

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
    MsgBox "Project1.Module1.CreateBitmapPicture" & Err & error
    Resume Next
End Function

