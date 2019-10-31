This is an example on how to capture Worksheet events through an Excel VBA Add-Ins (xlam file).

Introduction:
=============
Outside add-ins, it is usually quite easy to catch worksheet events in a VBA Macro. The IDE provides enough support and you don’t have to know so much fact about the events.
Unfortunately, it is not so easy when you create and use an Add-Ins. The code is executed in a different way, and it becomes more difficult to catch the worksheet events.
The example presents a way to catch the events on a worksheet. Following the state of the new selection, a customized toggle button will be activated or not. So, thanks this example you will know how to catch events in a worksheet through Add-Ins, but also know how to add new buttons in the Ribbon and how to deal with them.

External tools:
===============
To add new buttons in the Ribbon, you need to modify the XML file of your Add-ins (Excel files format is built around a zipped XML file). To do this, you can use Office Ribbon X Editor (available from https://github.com/fernandreu/office-ribbonx-editor) but other similar tools exist.
This tool will allow you to modify the XML content and loads image inside your Add-Ins (useful to customize the icons). It also provides a way to generate Callback functions (functions called when an action occurs on a button).


