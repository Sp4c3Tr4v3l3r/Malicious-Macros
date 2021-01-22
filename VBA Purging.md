## VBA Purging

- VBA you have 2 types of code
	- Source Code (CompressedSourceCode)
		- oledump -s 7s example-07.doc
		- oledump -s 7 -v example-07.doc
		- This is by design, it will always run
	- Compiler Code (PerformanceCache)
		- oledump -s 7c example-06.doc
		- It is version dependent and architecture dependent

- VBA Purging
	- Not keeping the compiler code (suppress)
		- You also need to update to dir (ModuleOFFSet1, ModuleOFFSet2)
		- This tell you where the compressed code starts
	- Keeping the source code
	
- [Module Stream: Visual Basic Modules](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/c66b58a6-f8ba-4141-9382-0612abce9926)



