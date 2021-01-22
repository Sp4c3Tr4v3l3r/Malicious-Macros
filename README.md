# Malicious Macros

## Intro to Macros and VBA

- VBA really allows a programmer to perform nearly any operation by a user.

## VBA Language for Script Kiddies

- VBA applications have access to a wide variety of system-level objects as well as the application and document objects. You can even reference external libraries, including DLLs, TLBs, OCXs

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

## Macro 4.0 (XLM Macros)

## Very Hidden Sheets

## VBA Stomping

## VBA Purging

## Code Signing Tampering

#### Disclaimer
*The content within this repository is for educational purposes only. It is designed to help users test their own system against information security threats and protect their IT infrastructure from similar attacks.*
