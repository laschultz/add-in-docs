
# DocumentSelectionChangedEventArgs.document property (JavaScript API for Office)
Gets a  **Document** object that represents the document that raised the **SelectionChanged** event.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Added in**|1.1|
[See all support details](#bk_support)




```
var myDoc = eventArgsObj.document;
```


## Return Value

A [Document](../reference/shared/document/document-object.md) object that represents the document that raised the[SelectionChanged](../reference/shared/document/selectionchanged-event/selectionchanged-event.md) event.


## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

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
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.0|Introduced|
