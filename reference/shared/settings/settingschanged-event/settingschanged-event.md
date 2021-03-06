
# Settings.settingsChanged event (JavaScript API for Office)
Occurs when the in-memory copy of the settings property bag is saved into the document with the [Settings.saveAsync](../reference/shared/settings/saveasync-method.md) method.

|||
|:-----|:-----|
|**Hosts:**|Excel, |
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Settings|
|**[Last changed](#bk_history) in**|1.0|
[See all support details](#bk_support)

```
Office.EventType.SettingsChanged
```


## Remarks

To add an event handler for the  **settingsChanged** event, use the[addHandlerAsync](../reference/shared/settings/addhandlerasync-method.md) method of the **Settings** object.

The  **settingsChanged** event fires only when your add-in's script calls the **Settings.saveAsync** method to persist the in-memory copy of the settings into the document file. The **settingsChanged** event is not triggered when the[Settings.set](../reference/shared/settings/set-method.md) or[Settings.remove](../reference/shared/settings/remove-method.md) methods are called.

The  **settingsChanged** event was designed to let you to handle potential conflicts when two or more users are attempting to save settings at the same time when your add-in is used in a shared (co-authored) document.


 **Important**  Your add-in's code can register a handler for the  **settingsChanged** event when the add-in is running with any Excel client, but the event will fire only when the add-in is loaded with a spreadsheet that is opened in Excel Online, _and_ more than one user is editing the spreadsheet (co-authoring). Therefore, effectively the **settingsChanged** event is supported only in Excel Online in co-authoring scenarios.


## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this event is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this event.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


|
|
||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||

|||
|:-----|:-----|
|**Available in requirement sets**|Settings|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>



|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|
