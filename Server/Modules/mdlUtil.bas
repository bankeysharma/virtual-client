Attribute VB_Name = "mdlUtil"
Option Explicit
Option Private Module

Private Declare Function GetComputerName _
    Lib "kernel32" _
    Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) _
    As Long

Public Sub Pause(ByVal nSecond As Single)
    Dim t0 As Single
    Dim dummy As Integer
    t0 = Timer
    Do While Timer - t0 < nSecond
    dummy = DoEvents()
    If Timer < t0 Then
    t0 = t0 - 24 * 60 * 60
    End If
    Loop
End Sub

Public Function ClientName() As String
    Dim Buffer As String
    Dim msg As String
    
    Buffer = Space(255)
    
    Call GetComputerName(Buffer, Len(Buffer))
    
    ClientName = left(Buffer, InStr(Buffer, " "))

    ClientName = StrConv(ClientName, vbUpperCase)
    
End Function


