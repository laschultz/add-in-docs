
# TableData object (JavaScript API for Office)
Represents the data in a table or a [TableBinding](../reference/shared/binding-object/tablebinding-object/tablebinding-object.md).

|||
|:-----|:-----|
|**Hosts:**|Excel, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|TableBindings|
|**[Added](#bk_history) in**|1.1|
[See all support details](#bk_support)

```
TableData
```

[![Try out this call in the interactive API Tutorial for Excel](../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.mdl?scenario=Set+Formatting)

## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[headers](../reference/shared/tabledata/headers-property.md)|Gets or sets the headers in the table.|
|[rows](../reference/shared/tabledata/rows-property.md)|Gets or sets the rows in the table.|

## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


||
|:-----|
|**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|TableBindings|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word Online.|
|1.1|Added support for Excel and Word in Office for iPad|
|1.0|Introduced|
