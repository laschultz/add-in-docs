
# Bindings object (JavaScript API for Office)
Represents the bindings the add-in has within the document.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**[Last changed](#bk_history)** in|1.1|
[See all support details](#bk_support)

```
Office.context.document.bindings
```


**Properties**

|||
|:-----|:-----|
|Name|Description|
|[document](../reference/shared/bindings-object/document-property.md)|Gets a  **Document** object that represents the document associated with this set of bindings.|

**Methods**

|||
|:-----|:-----|
|Name|Description|
|[addFromNamedItemAsync](../reference/shared/bindings-object/addfromnameditemasync-method.md)|Adds a binding to a named item in the document.|
|[addFromPromptAsync](../reference/shared/bindings-object/addfrompromptasync-method.md)|Displays UI that enables the user to specify a selection to bind to.|
|[addFromSelectionAsync](../reference/shared/bindings-object/addfromselectionasync-method.md)|Adds a binding object of the type specified bound to the current selection in the document.|
|[getAllAsync](../reference/shared/bindings-object/getallasync-method.md)|Gets all bindings that were previously created.|
|[getByIdAsync](../reference/shared/bindings-object/getbyidasync-method.md)|Gets the specified binding by its identifier.|
|[releaseByIdAsync](../reference/shared/bindings-object/releasebyidasync-method.md)|Removes the specified binding.|

## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


|||||
|:-----|:-----|:-----|:-----|
||Office for Windows desktop|Office Online(in browser)|Office for iPad|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad|
|1.1|For [addFromNamedItemAsync](../reference/shared/bindings-object/addfromnameditemasync-method.md), [addFromPromptAsync](../reference/shared/bindings-object/addfrompromptasync-method.md), and [addFromSelectionAsync](../reference/shared/bindings-object/addfromselectionasync-method.md) added support for binding to matrix data as a table binding in add-ins for Excel.|
|1.1|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>For <a href="8fa0cb4a-fad1-4f2e-9a7e-5f7aa7789eca.htm">document</a> property, added access to a <span class="keyword">Document</span> object that represents the current Access database in content add-ins for Access. </p></li><li><p>For all methods, added support for table binding in content add-ins for Access. </p></li></ul>|
|1.0|Introduced|
