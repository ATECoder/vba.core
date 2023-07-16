# About

[cc.isr.test.fx] provides a test framework for Visual Basic for Applications.

## Workbook references

* [cc.isr.core.io] - Core I/O workbook.

## Object Libraries references

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]

## Key Features

* A core class outcome for test functions.
* A test executive.
* An enumerator of test methods and modules.

# Main Types

The main types provided by this library are:

* _Assert_ - Returns results from unit tests.
* _TestExecutive_ - Singleton. A rudimentary unit test executive.
* _VbComponentExtensions_ - Singleton. Extension methods for the VBA VB Component object.

## Scripts

* _Deploy_: copies files to the build folder.
* _run.unit.tests.ps1_: a generic script for running unit tests.
* _cc.isr.test.fx.unit.tests_: a shortcut for running the unit tests.

## [Testing]

Unit testing can be accomplished using the power shell [Generic Test Script] which is inoked by the [Test Script shortcut]. 

Tests can also run by running the _Run Tests_ method from the _Testing_ worksheet.

# Feedback

[cc.isr.test.fx] is released as open source under the MIT license.
Bug reports and contributions are welcome at the [cc.isr.Core] repository.

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.Core.IO]: https://github.com/ATECoder/vba.core/io
[cc.isr.test.fx]: https://github.com/ATECoder/vba.core/src/testFx
[Test Script shortcut]: ./cc.isr.test.fx.unit.test
[Generic Test Script]: ./run.unit.tests.ps1

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>

