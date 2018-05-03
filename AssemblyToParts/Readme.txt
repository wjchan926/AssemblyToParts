	Assembly To Parts Addin Project
	===============================
	
	Summary:
	===============================
	This Add-in provides Marlin engineers wtih the ability to push parameters from the assembly file directly to all of it's child files.  It will also update the corresponding iLogic required between assembly and child files.  This add-in also provides the ability to pull parameters from a parent assembly file, if the user is currently creating a part or assembly child file.

	Requirements:
	=============
	Autodsk Inventor 2017 or later is required to run this addin.
	Support for earlier versions has not been tested.

	Installation:
	=============
	Installation files are not provided.
	Contact wesley.j.chan@gmail.com for installation details.

	License:
	========
	The MIT License (MIT)

	Copyright (c) 2018 - Wesley Chan

	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in all
	copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
	SOFTWARE.

	If installing on Autodesk Inventor 2011 and earlier:
	(Support not test, use at own risk)
	How to Register/Unregister 
	=======================

	1) Build Project;

	2) Copy add-in dll file to one of following locations: 
		a) Anywhere, then *.addin file <Assembly> setting should be updated to the full path including the dll name
		b) Inventor <InstallPath>\bin\ folder, then *.addin file <Assembly> setting should be the dll name only: <AddInName>.dll
		c) Inventor <InstallPath>\bin\XX folder, then *.addin file <Assembly> setting shoule be a relative path: XX\<AddInName>.dll

	3) Copy.addin manifest file to one of following locations:
		a) Inventor Version Dependent
		Windows XP:
			C:\Documents and Settings\All Users\Application Data\Autodesk\Inventor [version]\Addins\
		Windows7/Vista:
			C:\ProgramData\Autodesk\Inventor [version]\Addins\

		b) Inventor Version Independent
		Windows XP:
			C:\Documents and Settings\All Users\Application Data\Autodesk\Inventor Addins\
		Windows7/Vista:
			C:\ProgramData\Autodesk\Inventor Addins\

		c) Per User Override
		Windows XP:
			C:\Documents and Settings\<user>\Application Data\Autodesk\Inventor [version]\Addins\
		Windows7/Vista:
			C:\Users\<user>\AppData\Roaming\Autodesk\Inventor [version]\Addins\

	4) Startup Inventor, the AddIn should be loaded

	To unregister the AddIn, remove the Autodesk.<AddInName>.Inventor.addin from above mentioned .addin manifest file locations directly.
