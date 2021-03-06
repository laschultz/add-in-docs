
# TextBinding object (JavaScript API for Office)
Represents a bound text selection in the document.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, PowerPoint, Project, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|TextBindings|
|**[Added](#bk_history) in**|1.0|
[See all support details](#bk_support)

```
TextBinding
```


## Remarks

The  **TextBinding** object inherits the[id](../reference/shared/binding-object/id-property.md) property,[type](../reference/shared/binding-object/type-property.md) property,[getDataAsync](../reference/shared/binding-object/getdataasync-method.md) method, and[setDataAsync](../reference/shared/binding-object/setdataasync-method.md) method from the[Binding](../reference/shared/binding-object/binding-object.md) object. It does not implement any additional properties or methods of its own.


## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|TextBindings|
|**Minimum permission level**|[WriteDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad.|
|1.0|Introduced|
