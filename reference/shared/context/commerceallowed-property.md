
# Context.commerceAllowed property (JavaScript API for Office)
Gets whether the add-in is running on a platform that allows links to external payment systems.

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**[Last changed](#bk_history) in**|1.1|

```
var allowCommerce = Office.context.commerceAllowed;
```


## Return value

 ****Returns  **True** if developers can display sell or upgrade UI in the add-in on that platform; otherwise returns **False**.


## Remarks

The iOS App Store doesn't support apps with add-ins that provide links to additional payment systems. However, Office Add-ins running on the Windows desktop or for Office Online in the browser, do allow such links. If you want the UI of your add-in to provide a link to an external payment system on platforms other than iOS, you can use the  **commerceAllowed** property to control when that link is displayed.


## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


|
|
||**Office for iPad**|
|:-----|:-----|
|**Excel**|Y|
|**PowerPoint**||
|**Word**|Y|

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
|1.1|Introduced.|
