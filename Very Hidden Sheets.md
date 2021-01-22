## Very Hidden Sheets

- Hidden Macro 4.0
	- oledump -p pluggin_biff --pluginoptions "-o BOUNDSHEET" example-03.xls

- Very Hidden Macro 4.0
	- [BoundSheet8](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/b9ec509a-235d-424e-871d-f8e721106501)
	- You can replace it in a HexaDecimal Editor
	- 0x00 - Visible.
	- 0x01 - Hidden.
	- 0x02 - Very Hidden; the sheet (1) is hidden and cannot be displayed using the user interface.
	- 0x03 - Very, Very Hidden? Protected View

- Unused bits
	- unused (6 bits): Undefined and MUST be ignored.
	- Change the very hidden value, and a value of the unused 6 bits. Now you don't have the Protected View
	- oledump -p pluggin_biff --pluginoptions "-o BOUNDSHEET -a" example-03.xls
	
- This way you can avoid YARA Rules
