Attribute VB_Name = "Connection"
Public flag As Integer
Dim Disk As String
Public conn As ADODB.Connection
Public rsPay As ADODB.Recordset
Public rsExpiry As ADODB.Recordset
Public rsMem As ADODB.Recordset
Public rsAcc As ADODB.Recordset
Public rsDef As ADODB.Recordset
Public rsComm As ADODB.Recordset
Public rsRefIdCheck As ADODB.Recordset
Public Sub Main()
On Error GoTo merr
flag = 0
chpass = 0
Dim str1 As String

Set conn = New ADODB.Connection
'str1 = "DSN=MS Access Database;DBQ=C:\Sardar\Data.mdb;DefaultDir=C:\Sardar;DriverId=25;FIL=MS Access;FILEDSN=C:\Program Files\Common Files\ODBC\Data Sources\MS Access Database (not sharable).dsn;MaxBufferSize=2048;PageTimeout=5;"
str1 = "provider=microsoft.jet.oledb.3.51;Jet OLEDB:database password=barcode420;data source="
str1 = str1 & App.Path & "\data\data97.mdb"
conn.Open str1

On Error GoTo off2000
GoTo connected
off2000:
' This is used for Office 2000
str1 = "provider=microsoft.jet.oledb.4.0;Jet OLEDB:database password=master;data source="
str1 = str1 & App.Path & "\data\data97.mdb"

'This is used for Office 97
'str1 = "provider=microsoft.jet.oledb.3.51;Jet OLEDB:database password=barcode420;data source="
connected:

Set rsMem = New ADODB.Recordset
rsMem.Open "select * from MastCust", conn, adOpenStatic, adLockOptimistic

Set rsRefIdCheck = New ADODB.Recordset
rsRefIdCheck.Open "select * from MastCust", conn, adOpenStatic, adLockOptimistic


Set rsAcc = New ADODB.Recordset
rsAcc.Open "select * from Account", conn, adOpenStatic, adLockOptimistic

Set rsPay = New ADODB.Recordset
rsPay.Open "select * from PayList", conn, adOpenStatic, adLockOptimistic


Set rsComm = New ADODB.Recordset
rsComm.Open "select * from MastComm", conn, adOpenStatic, adLockOptimistic


Set rsDef = New ADODB.Recordset
rsDef.Open "MastDefault", conn, adOpenStatic, adLockOptimistic


Set rsExpiry = New ADODB.Recordset
rsExpiry.Open "select * from MastExpiry", conn, adOpenStatic, adLockOptimistic

Disk = "C:\"
'*****************************************************************************************************************
' This will set the expiry date of the software when it is installed for 1st time
' At the time of Installation
If rsExpiry.RecordCount = 0 Then
    rsExpiry.AddNew
    If rsExpiry!usage = 0 Then
        rsExpiry!InstallDate = Date
        rsExpiry!expirydate = DateAdd("d", 30, Date)
        rsExpiry!Paid = "N"
        rsExpiry!CurrentYear = Date
        rsExpiry!YearChanged = "N"
        rsExpiry!HDSerial = VolumeSerialNumber(Disk)
    End If
    rsExpiry.Update
End If


If rsExpiry!HDSerial <> VolumeSerialNumber(Disk) Then
    MsgBox "Contact Developer for Re-installing the Software !                                             Developed by: AnaSys Softwares (Ahmedabad), Vikrant Thakker,                                      Ph : 7911226, 7911833, Email : vikrant_thakker@hotmail.com", vbCritical, "Developer"
    End
End If

    rsExpiry!usage = rsExpiry!usage + 1
    rsExpiry.Update


'******************************************************************************************************************
' If user has not paid then This will check for the expiry date and usage, everytime the program is runned...
' After installation
If rsExpiry!Paid = "N" Then
    If rsExpiry!expirydate = Date Or rsExpiry!expirydate < Date Or rsExpiry!usage > 100 Then
        rsExpiry!accexpired = True
        rsExpiry.Update
    End If
End If

'******************************************************************************************************************
' After Trial version has expired
If rsExpiry!Paid = "N" Then

    If rsExpiry!accexpired = True Then
        'MsgBox "Trial Version Expired ! Contact Vikrant Thakker  for Final Version.     phone:(079)7911226,7911833, email : vikrant_thakker@hotmail.com", vbCritical
        MsgBox "Kindly Pay the negotiated amount for further usage of this Software !                          Developed by: AnaSys Softwares (Ahmedabad), Vikrant Thakker, Ph : 7911226, 7911833, Email : vikrant_thakker@hotmail.com", vbCritical, "Developer"
        
        End
    End If
End If

If Year(Now) < Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in previous working year ! Contact Developer", vbCritical, "MLM"
    rsExpiry!YearChanged = "Y"
    rsExpiry.Update
End
ElseIf Year(Now) > Year(rsExpiry!CurrentYear) Then
    MsgBox "You cannot work in next working year ! Contact Developer", vbCritical, "MLM"
    rsExpiry!YearChanged = "Y"
    rsExpiry.Update
End
End If

' Once the year gets changed then even if the user changes the year of his computer
'manually, he should not be allowed to use the software
' The software can be used only after the developer writes the status of the YearChanged="N" in the database
' Also the developer needs to write the new currentyear

If rsExpiry!YearChanged = "Y" Then
    MsgBox "You cannot work in changed working year ! Contact Developer !", vbCritical, "MLM"
    End
End If

'If Year(Now) <> rsExpiry!CurrentYear Then
'    MsgBox "Please call developer to run the software in Changed Year"
'    End
'End If


'rsExpiry!expirydate = DateAdd("d", 30, Date)



'Load frmMain
'frmMain.Show
Load frmSplash
frmSplash.Show
Exit Sub
merr:
MsgBox Err.Description, vbOKOnly, "MLM"
End Sub

