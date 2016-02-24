
# Shared API


The Shared API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in all three types of Office Add-ins: content, task pane, and Outlook add-ins.


## Objects





|**Object**|**Description**|
|:-----|:-----|
|[AsyncResult](../reference/shared/asyncresult-object.md)|An object which encapsulates the result of an asynchronous request, including status and error information if the request failed.|
|[Context](../reference/shared/context/context-object.md)|Represents the runtime environment of the add-in and provides access to key objects of the API.|
|[Error](../reference/shared/error/error-object.md)|Provides specific information about an error that occurred during an asynchronous data operation.|
|[Office](../reference/shared/office/office-object.md)|Represents an instance of an add-in, which provides access to the top-level objects of the API.|


|**Member**|**Description**|
|:-----|:-----|
|[event.completed](../reference/shared/office/event.completed.md)|The callback that the add-in invokes to let Outlook know that the operation is done.|
|[event.source.id](../reference/shared/office/event.source.id.md)|Gets the id of the control that triggered calling this function.|

## Supported host applications


|||
|:-----|:-----|
|**Supported hosts**|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Access</p></li><li><p>Excel</p></li><li><p>Outlook</p></li><li><p>PowerPoint</p></li><li><p>Project</p></li><li><p>Word</p></li></ul>|
|**Library**|Office.js|
|**Namespace**|Office|

## Additional resources
<a name="bk_addresources"> </a>


- [JavaScript API for Office](../reference/javascript-api-for-office.md)
    
