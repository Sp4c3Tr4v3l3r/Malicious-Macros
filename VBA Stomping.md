## VBA Stomping

- Add simple VBA script
	- oledump -i example-06.doc
	
- VBA you have 2 types of code
	- Source Code (CompressedSourceCode)
		- oledump -s 7s example-06.doc
		- oledump -s 7 -v example-06.doc
	- Compiler Code (PerformanceCache)
		- oledump -s 7c example-06.doc
		- It is version dependent and architecture dependent

- VBA Stomping
	- Just keeping the compiler code
	- Not keeping the source code (alter or suppress)
	
- [Module Stream: Visual Basic Modules](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/c66b58a6-f8ba-4141-9382-0612abce9926)
