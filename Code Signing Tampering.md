## Code Signing Tampering

- VBA you have 2 types of code
	- Source Code (CompressedSourceCode)
		- oledump -s 7s example-07.doc
		- oledump -s 7 -v example-07.doc
		- This is by design, it will always run
	- Compiler Code (PerformanceCache)
		- oledump -s 7c example-06.doc
		- It is version dependent and architecture dependent

- Code Signing Tampering
	- Compiler code (alter)
		- It is ignored for content hashes
		- Launch calc
	- Source Code
		- For digital signatures in VBA Project by design Microsoft only takes into account Source Code.
		- MsgBox Hello (signed)
		
- Take an existing document, that is signed.
	- oledump -s a3 -v example-08.docm
	- pcodedmp.exe example-08.docm | tail
	
- [Contents Hashes](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/73422f49-565e-47a3-baf4-5742e2ba7dad)



