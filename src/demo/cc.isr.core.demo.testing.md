# Testing the [cc.isr.Core] Workbook

[cc.isr.Core.demo] is an Excel workbook for demonstrating some functionality of the [cc.isr.core] workbook.

## Worksheets

The [cc.isr.Core.Demo] workbook includes two worksheets: 

* Countdown Timer -- To test the `EventTimer` class.

## Integration Testing

### Testing the EventTimer using the Countdown Timer Worksheet

Use the following procedure to run the `EventTimer` tests:

1) Select the _Countdown Timer_ Worksheet;
2) Click ___Reset Timer___ button to initialize the timer duration to 15;
3) Click ___Start Timer___ to start the countdown. after a short pause, the display with decrement at fractions of a second intervals;
4) Click ___Stop Timer___ to pause the timer;
5) Click ___Dispose___ to stop and terminate the timer.
6) Close the Excel workbook and check the task manager to ensure that all Excel instances are closes.

To validate item 6, make sure to start the test session with only a single Excel instance.

[cc.isr.Core]: https://github.com/ATECoder/vba.core
[cc.isr.Core.demo]: https://github.com/ATECoder/vba.core/src/demo

[ISR]: https://www.integratedscientificresources.com

[Microsoft Scripting Runtime]: c:\windows\system32\scrrun.dll
[Microsoft Visual Basic for Applications Extensibility 5.3]: <c:/program&#32;files/common&#32;files/microsoft&#32;shared/vba/vba7.1/vbeui.dll>
