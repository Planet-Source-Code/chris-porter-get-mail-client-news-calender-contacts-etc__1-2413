<div align="center">

## Get mail client, news, calender, contacts, etc\.


</div>

### Description

Finds the default client (default program which Internet Explorer uses) for mail, news, contacts, calender, internet call.
 
### More Info
 
Which client you want to find where is stored, for example mail, news or contacts.

Has only been tested with IE 5.0, and 4.0 under Win98. It might run on Win95, and IE versions lower than 5.0.

It returns where the client is stored on the hard drive.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Porter](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-porter.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-porter-get-mail-client-news-calender-contacts-etc__1-2413/archive/master.zip)

### API Declarations

```
'All is pasted into the form.
```


### Source Code

```
Option Explicit
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Type TypesOfClient
Mail As String
News As String
Calendar As String
Contacts As String
Internet_Call As String
End Type
'Get the registry keys for the programs location
Function GetReg(hInKey As Long, ByVal subkey As String, ByVal valname As String)
Dim RetVal As String, hSubKey As Long, dwType As Long
Dim SZ As Long, v As String, r As Long
RetVal = ""
r = RegOpenKeyEx(hInKey, subkey, 0, 983139, hSubKey)
If r <> 0 Then GoTo Ender
SZ = 256: v = String(SZ, 0)
r = RegQueryValueEx(hSubKey, valname, 0, dwType, ByVal v, SZ)
If r = 0 And dwType = 1 Then
RetVal = Left(v$, SZ - 1)
Else
RetVal = ""
End If
If hInKey = 0 Then r = RegCloseKey(hSubKey)
Ender:
GetReg = RetVal
End Function
Private Function GetClient() As TypesOfClient
Static KeyName As String, O(5) As String, i As Byte, d As String
O(1) = "Mail"
O(2) = "News"
O(3) = "Calendar"
O(4) = "Contacts"
O(5) = "Internet Call"
'In this tedious method I have to get all 5.
For i = 1 To 5
KeyName = "Software\Clients\" + O(i) + "\"
d = GetReg(&H80000002, KeyName, "")
KeyName = KeyName + d + "\Shell\Open\Command\"
d = GetReg(&H80000002, KeyName, "")
O(i) = d
Next i
'Set the values to where the programs were found.
GetClient.Mail = O(1)
GetClient.News = O(2)
GetClient.Calendar = O(3)
GetClient.Contacts = O(4)
GetClient.Internet_Call = O(5)
End Function
Private Sub Form_Load()
'Run the mail client
Shell GetClient.Mail
End Sub
```

