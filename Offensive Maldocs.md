## Offensive Maldocs

- Macro 4.0 (XLM Macros)
	- No. This is not VBA.
	- No. This is not XML.
	- XLM (Excel 4.0) Macros were created in 1992, before VBA existed.

	- Background
		- [Old School: Evil Excel 4.0 Macros](https://outflank.nl/blog/2018/10/06/old-school-evil-excel-4-0-macros-xlm/)
		- [Excel4-DCOM](https://github.com/outflanknl/Excel4-DCOM]
		- [Process Injection with SYLK](https://outflank.nl/blog/2019/10/30/abusing-the-sylk-file-format/)
		- [SharpShooter](https://github.com/mdsecactivebreach/SharpShooter)
		- [Cybereason 64-Bit Shellcode Execution via Excel 4.0](https://www.cybereason.com/blog/excel4.0-macros-now-with-twice-the-bits)
		- [Macrome](https://github.com/michaelweber/Macrome)
		- [EXCELntDonut](https://github.com/FortyNorthSecurity/EXCELntDonut)

	- Why is this so great?
		- AMSI has no vibility into XML macros.
		- A/V still not great at detecting malicious XLM macros.
		- Payloads can be delivered in .xls files.
		- Process injection is possible.

	- Examples
		- [Living of the land](https://inquest.net/blog/2019/01/29/Carving-Sneaky-XLM-Files)
		- [XLM Macros 101](https://hatching.io/blog/excel-xlm-extraction/)

	- Checkout
		- [Donut](https://github.com/TheWover/donut)
			-Take a dotnet assembly and convert that into independent shellcode.
		- [Further Evasion in the Forgotten Corners of MS-XLS](https://malware.pizza/2020/06/19/further-evasion-in-the-forgotten-corners-of-ms-xls/)
		- [Evading Detection with Excel 4.0 Macros and the BIFF8 XLS Format](https://malware.pizza/2020/05/12/evading-av-with-excel-macros-and-biff8-xls/)

- A/V Evasion
	- Char Functions
	- Hide Macro Sheet
	
- [Epic Manchego](https://blog.nviso.eu/2020/09/01/epic-manchego-atypical-maldoc-delivery-brings-flurry-of-infostealers/)

- [Hot Manchego](https://github.com/FortyNorthSecurity/hot-manchego)

- [Hover With Power](https://github.com/ethanhunnt/Hover_with_Power)
