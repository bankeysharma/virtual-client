Attribute VB_Name = "mdlIPAddresses"
Option Explicit
 
Private Const ERROR_SUCCESS As Long = 0

Private Type MIB_IPADDRROW
  dwAddr As Long        'IP address
  dwIndex As Long       'index of interface associated with this IP
  dwMask As Long        'subnet mask for the IP address
  dwBCastAddr As Long   'broadcast address (typically the IP
                        'with host portion set to either all
                        'zeros or all ones)
  dwReasmSize As Long   'reassembly size for received datagrams
  unused1 As Integer    'not currently used (but shown anyway)
  unused2 As Integer    'not currently used (but shown anyway)
End Type

Private Declare Function GetIpAddrTable Lib "iphlpapi.dll" _
  (ByRef ipAddrTable As Byte, _
   ByRef dwSize As Long, _
   ByVal bOrder As Long) As Long
   
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (dst As Any, src As Any, ByVal bcount As Long)
  
Private Declare Function inet_ntoa Lib "wsock32" _
   (ByVal addr As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, _
   ByVal Ptr As Long) As Long
                        
Private Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long
   

Private Sub Form_Load()

'   With ListView1
'      .View = lvwReport
'      .ColumnHeaders.Add , , "Index"
'      .ColumnHeaders.Add , , "IP Address"
'      .ColumnHeaders.Add , , "Subnet Mask"
'      .ColumnHeaders.Add , , "Broadcast Addr"
'      .ColumnHeaders.Add , , "Reassembly"
'      .ColumnHeaders.Add , , "unused1"
'      .ColumnHeaders.Add , , "unused2"
'   End With
   
End Sub

      
Public Function GetInetStrFromPtr(ByVal Address As Long) As String
  
   GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))

End Function


Public Function GetStrFromPtrA(ByVal lpszA As Long) As String

   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
   
End Function


Public Sub updateIPAddressTable()
   
   Dim IPAddrRow As MIB_IPADDRROW
   Dim buff() As Byte
   Dim cbRequired As Long
   Dim nStructSize As Long
   Dim nRows As Long
   Dim cnt As Long
   Dim itmx As ListItem
   
   Call GetIpAddrTable(ByVal 0&, cbRequired, 1)

   If cbRequired > 0 Then
    
      ReDim buff(0 To cbRequired - 1) As Byte
      
      If GetIpAddrTable(buff(0), cbRequired, 1) = ERROR_SUCCESS Then
      
        'saves using LenB in the CopyMemory calls below
         nStructSize = LenB(IPAddrRow)
   
        'first 4 bytes is a long indicating the
        'number of entries in the table
         CopyMemory nRows, buff(0), 4
      
        '**********************************************
        'Setting Progress bar
        '**********************************************
        Call setPrgBar("Updating IPAddresses", 1, nRows)
        
        For cnt = 1 To nRows
         
           'moving past the four bytes obtained
           'above, get one chunk of data and cast
           'into an IPAddrRow type
           prgBar.Value = prgBar.Value + 1
           DoEvents
           
            CopyMemory IPAddrRow, buff(4 + (cnt - 1) * nStructSize), nStructSize
            
           'pass the results to the listview
            With IPAddrRow
                
                orsIPAddresses.Filter = "IPAddress = '" & _
                                        GetInetStrFromPtr(.dwAddr) & "' "
                
                If (orsIPAddresses.BOF And orsIPAddresses.EOF) Then
                    Call orsIPAddresses.AddNew(Array("IPAddress"), _
                                Array(GetInetStrFromPtr(.dwAddr)))
                End If
                
            End With

        Next cnt
        
        '*****************************************
        'Closing Progress bar
        '*****************************************
        Call offPrgBar
        
        orsIPAddresses.Filter = adFilterNone
        
      End If
   End If
   
End Sub

'Private Sub Command1_Click()
'
'   Dim IPAddrRow As MIB_IPADDRROW
'   Dim buff() As Byte
'   Dim cbRequired As Long
'   Dim nStructSize As Long
'   Dim nRows As Long
'   Dim cnt As Long
'   Dim itmx As ListItem
'
'   Call GetIpAddrTable(ByVal 0&, cbRequired, 1)
'
'   If cbRequired > 0 Then
'
'      ReDim buff(0 To cbRequired - 1) As Byte
'
'      If GetIpAddrTable(buff(0), cbRequired, 1) = ERROR_SUCCESS Then
'
'        'saves using LenB in the CopyMemory calls below
'         nStructSize = LenB(IPAddrRow)
'
'        'first 4 bytes is a long indicating the
'        'number of entries in the table
'         CopyMemory nRows, buff(0), 4
'
'         For cnt = 1 To nRows
'
'           'moving past the four bytes obtained
'           'above, get one chunk of data and cast
'           'into an IPAddrRow type
'            CopyMemory IPAddrRow, buff(4 + (cnt - 1) * nStructSize), nStructSize
'
'           'pass the results to the listview
''            With IPAddrRow
''                Set itmx = ListView1.ListItems.Add(, , GetInetStrFromPtr(.dwIndex))
''                itmx.SubItems(1) = GetInetStrFromPtr(.dwAddr)
''                itmx.SubItems(2) = GetInetStrFromPtr(.dwMask)
''                itmx.SubItems(3) = GetInetStrFromPtr(.dwBCastAddr)
''                itmx.SubItems(4) = GetInetStrFromPtr(.dwReasmSize)
''                itmx.SubItems(5) = (.unused1)
''                itmx.SubItems(6) = (.unused2)
''            End With
''
'          Next cnt
'
'      End If
'   End If
'
'End Sub
'--end block--'




