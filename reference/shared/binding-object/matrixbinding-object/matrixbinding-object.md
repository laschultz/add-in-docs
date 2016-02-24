
# MatrixBinding object (JavaScript API for Office)
Represents a binding in two dimensions of rows and columns. 

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|MatrixBindings|
|**[Last changed](#bk_history) in Selection**|1.1|
[See all support details](#bk_support)

```
MatrixBinding
```


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[columnCount](../reference/shared/binding-object/matrixbinding-object/columncount-property.md)|Gets the number of columns in the matrix data structure, as an integer value.|
|[rowCount](../reference/shared/binding-object/matrixbinding-object/rowcount-property.md)|Gets the number of rows in the matrix data structure, as an integer value.|

## Remarks

The  **MatrixBinding** object inherits the[id](../reference/shared/binding-object/id-property.md) property,[type](../reference/shared/binding-object/type-property.md) property,[getDataAsync](../reference/shared/binding-object/getdataasync-method.md) method, and[setDataAsync](../reference/shared/binding-object/setdataasync-method.md) method from the[Binding](../reference/shared/binding-object/binding-object.md) object.


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
|**Available in requirement sets**|MatrixBindings|
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
