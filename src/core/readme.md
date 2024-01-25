# About

[cc.isr.Core] is an Excel workbook with core Visual Basic for Applications modules and classes that support [ISR] workbooks.

## Workbook references

* [cc.isr.core.io] - Core I/O workbook.

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]
* [Microsoft VBScript Regular Expression 5.5]

## Key Features

* Provides core classes such as `EventTimer` and 'StringBuilder'.
* Provide extension classes such as `StringExtensions` and`PathExtensions`.
* Provides a rudimentary test executive.

# Main Types

The main types provided by this library are:

* _Assert_ - Returns results from unit tests.
* _CanceEventArg_ - Event arguments for canceling event handlers.
* _CollectionExtensions_ - Singleton. Collection extensions.
* _MacroInfo_ - Holds information such as name and module name about Excel Macro methods.
* _Marshal_ - Singleton. Supports Endianess.
* _ModuleInfo_ - Holds information such as name and project name about Excel modules.
* _EventTimer_ - A timer class capable of issuing events with millisecond time resolution.
* _StopWatch_ - A high resolution stop watch using the Windows API.
* _StringBuilder_ - A fast string builder.
* _StringExtensions_ - Singleton. String extensions.
* _TestExecutive_ - Singleton. A rudimentary unit test executive.
* _UserDefinedError_ - A user defined error class.
* _UserDefinedErrors_ - Manages the user defined errors.
* _WorkbookUnilities_ - Singleton. Enumerates test methods.

## Scripts

* _Deploy_: copies files to the build folder.

## [Testing]

Testing can be accomplished using the [cc.isr.core.test] workbook.

## [User-Defined Type Not Defined error]

Occasionally, this error message displays when compiling this project. This was solved as follows:
1. save a copy;
2. if this does not work:
	1. Brows for the workbook in Excel;
	2. From the _Open_ menu dropdown select the repare option.

# Feedback

[cc.isr.Core] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.Core] repository.

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.core.io]: https://github.com/ATECoder/vba.core/src/io
[cc.isr.core.test]: https://github.com/ATECoder/vba.core/src/test

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
* [Microsoft VBScript Regular Expression 5.5]: <c:/windows/system32/vbscript.dll/3>
[User-Defined Type Not Defined error]: https://stackoverflow.com/questions/19680402/compile-throws-a-user-defined-type-not-defined-error-but-does-not-go-to-the-of

