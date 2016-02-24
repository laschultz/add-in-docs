
# SettingsChangedEventArgs object (JavaScript API for Office)
Provides information about the settings that raised the [settingsChanged](../reference/shared/settings/settingschanged-event/settingschanged-event.md) event.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, |
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Settings|
|**[Last changed](#bk_history) in**|1.0|
[See all support details](#bk_support)

```
Office.EventType.SettingsChanged
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[settings](../reference/shared/settings/settingschanged-event/args/settings-property.md)|Gets a  **Settings** object that represents the settings that raised the **settingsChanged** event.|
|[type](../reference/shared/settings/settingschanged-event/args/type-property.md)|Get an  **EventType** enumeration value that identifies the kind of event that was raised.|

## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

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
