Attribute VB_Name = "GetSerialNo"
'Call this procedure like this:
'Serial=VolumeSerialNumber("C:\")
Public Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function VolumeSerialNumber(ByVal RootPath As String) As String
    
    Dim VolLabel As String
    Dim VolSize As Long
    Dim Serial As Long
    Dim MaxLen As Long
    Dim Flags As Long
    Dim Name As String
    Dim NameSize As Long
    Dim S As String

    If GetVolumeSerialNumber(RootPath, VolLabel, VolSize, Serial, MaxLen, Flags, Name, NameSize) Then
        'Create an 8 character string
        S = Format(Hex(Serial), "00000000")
        'Adds the '-' between the first 4 characters and the last 4 characters
        VolumeSerialNumber = Left(S, 4) + "-" + Right(S, 4)
    Else
        'If the call to API function fails the function returns a zero serial number
        VolumeSerialNumber = "0000-0000"
    End If

End Function



