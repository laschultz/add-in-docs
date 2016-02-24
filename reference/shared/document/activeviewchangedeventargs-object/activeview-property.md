
# DocumentActiveViewChangedEventArgs.activeView property (JavaScript API for Office)
Gets an  **ActiveView** enumeration value that identifies the state of the active view of the document, for example, whether the user can edit the document.

|||
|:-----|:-----|
|**Hosts:**|PowerPoint|
|**[Added](#bk_history) in**|1.1|
[See all support details](#bk_support)

```
var myView = eventArgsObj.activeView;
```


## Return Value

The [ActiveView](../reference/enumerations/activeview-enumeration.md) of the view that raised the event.


## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**PowerPoint**|Y|Y|Y|

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for PowerPoint in Office for iPad.|
|1.1|Introduced|
