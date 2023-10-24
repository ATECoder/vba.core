# Testing the [cc.isr.Core] and [cc.isr.Core.IO] Workbooks.

[cc.isr.Core.Test] is an Excel workbook for testing the [cc.isr.core] workbook.

## Worksheets

The [cc.isr.Core.Test] workbook includes the following worksheet: 

* TestSheet -- To run unit tests.

## Unit Testing

To enable unit testing, the Excel _Trust Center_, which can be found from the _Search_ box, 
and check _Trust access to the VBA project object model_ from the _Macro Settings_ in the _Trust Center_.  

### Unit testing command link

Clicking the [Unit Test Link] file runs the unit tests.

### Unit testing with the TestSheet Worksheet

Use the following procedure to run unit tests:
1) Click the ___List Tests___ button.
2) The drop down list now includes the list of available test suites;
3) Select a test from the list;
4) Click ___Run Selected Tests___;
   * The list of tests included in the test suite will display.
   * Passed tests display Passed with a green background;
   * Failed tests display Fail with a red background and a message describing the failure.

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.Core.IO]: https://github.com/ATECoder/vba.core/src/io
[cc.isr.Core.Test]: https://github.com/ATECoder/vba.core/src/test

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>

[Unit Test Link]: ./cc.isr.core.test.unit.test.lnk