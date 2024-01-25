# VBA Core Demo

[cc.isr.Core.demo] Excel workbook demonstrates some functionality of the [cc.isr.Core] workbook classes.

## Dependencies

The [cc.isr.core.demo] workbook depends on the following Workbooks:

* [cc.isr.Core] - Includes core Visual Basic for Applications classes and modules.

## References

The following object libraries are used as references:

* [Microsoft Scripting Runtime]
* [Microsoft Visual Basic for Applications Extensibility 5.3]

## Worksheets

The [cc.isr.Core] demo workbook includes two worksheets: 

* TestSheet -- To run unit tests.
* Countdown Timer -- To test the `EventTimer` class.

## [Testing]

Testing information is included in the [Testing] document.

## Scripts

* [Deploy]: copies the main workbook and its referenced workbooks to the deploy folder.
* [Localize]: sets the folders of the referenced workbook of each workbook to the same folder as the  referencing workbook.

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.Core.Demo]: https://github.com/ATECoder/vba.core.demo
[Testing]: ./cc.isr.core.demo.testing.md
[Deploy]: ./deploy.ps1
[Localize]: ./localize.ps1

[ISR]: https://www.integratedscientificresources.com

