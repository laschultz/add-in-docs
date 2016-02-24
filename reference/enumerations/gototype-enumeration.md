
# GoToType enumeration (JavaScript API for Office)
Specifies the type of place or object to navigate to.

|||
|:-----|:-----|
|**Hosts:**|Excel, PowerPoint, Word|
|**[Added](#bk_history) in**|1.1|
[See all support details](#bk_support)

```
Office.GoToType
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|**Supported clients**|
|:-----|:-----|:-----|:-----|
|Office.GoToType.Binding|"binding"|Goes to a binding object using the specified binding id.|ExcelWord|
|Office.GoToType.NamedItem|"namedItem"|Goes to an item using that item's name, such as the name assigned to a table or range.In Excel, you can use any structured reference for a named range or table: "Worksheet2!Table1"|Excel|
|Office.GoToType.Slide|"slide"|Goes to a slide using the specified id.|PowerPoint|
|Office.GoToType.Index|"index"|Goes to the specified index by slide number or enum: **Office.Index.First** **Office.Index.Last** **Office.Index.Next** **Office.Index.Previous**|PowerPoint|

## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**PowerPoint**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Introduced|
