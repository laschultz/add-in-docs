
# CustomXMLNodeType enumeration (JavaScript API for Office)
Specifies the node type.

[See all support details](#bk_support)


|||
|:-----|:-----|
|**Hosts:**|Word|
|**[Last changed](#bk_history) in**|1.1|

[See all support details](#bk_support)


```
Office.CustomXMLNodeType
```


## Members


**Values**


|**Enumeration**|**Value**|**Description**|
|:-----|:-----|:-----|
|Office.CustomXMLNodeType.Attribute|"attribute"|The node is an attribute.|
|Office.CustomXMLNodeType.CData|"CData"|The node is a CData type.|
|Office.CustomXMLNodeType.NodeComment|"comment"|The node is a comment.|
|Office.CustomXMLNodeType.Element|"element"|The node is an element.|
|Office.CustomXMLNodeType.NodeDocument|"nodeDocument"|The node is a Document element.|
|Office.CustomXMLNodeType.ProcessingInstruction|"processingInstruction"|The node is a processing instruction.|
|Office.CustomXMLNodeType.Text|"text"|The node is a text node.|

## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this enumeration is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this enumeration.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Word in Office for iPad.|
|1.0|Introduced|
