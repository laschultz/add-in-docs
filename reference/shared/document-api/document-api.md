
# Document API


The Document API subset of the JavaScript API for Office includes objects, methods, properties and events that you can use in the two types of Office Add-ins associated with documents: content and task pane add-ins.


## Objects





|**Object**|**Description**|**Supported host applications**|
|:-----|:-----|:-----|
|[Binding](../reference/shared/binding-object/binding-object.md)|An abstract class that represents a binding to a section of the document.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Access</p></li><li><p>Excel</p></li><li><p>Word</p></li></ul>|
|[Bindings](../reference/shared/bindings-object/bindings-object.md)|Represents the bindings the add-in has within the document.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Access</p></li><li><p>Excel</p></li><li><p>Word</p></li></ul>|
|[CustomXmlNode](../reference/shared/customxmlnode-object/customxmlnode-object.md)|Represents an XML node in a tree in a document.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Word</p></li></ul>|
|[CustomXmlPart](../reference/shared/customxmlpart-object/customxmlpart-object.md)|Represents a single  **CustomXMLPart** in a **CustomXMLParts** collection.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Word</p></li></ul>|
|[CustomXmlParts](../reference/shared/customxmlparts-object/customxmlparts-object.md)|Represents a collection of  **CustomXMLPart** objects.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Word</p></li></ul>|
|[CustomXmlPrefixMappings](../reference/shared/customxmlprefixmappings-object/customxmlprefixmappings-object.md)|Represents a collection of custom namespace prefix mappings.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Word</p></li></ul>|
|[Document](../reference/shared/document/document-object.md)|An abstract class that represents the document the add-in is interacting with.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Access</p></li><li><p>Excel</p></li><li><p>PowerPoint</p></li><li><p>Project</p></li><li><p>Word</p></li></ul>|
|[File](../reference/shared/file/file-object.md)|Represents the document file associated with an Office Add-in.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>PowerPoint</p></li><li><p>Word</p></li></ul>|
|[MatrixBinding](../reference/shared/binding-object/matrixbinding-object/matrixbinding-object.md)|Represents a binding in two dimensions of rows and columns. |
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Excel</p></li><li><p>Word</p></li></ul>|
|[ProjectDocument](../reference/shared/projectdocument/projectdocument-object.md)|An abstract class that represents the project document (the active project) with which the Office Add-in interacts.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Project</p></li></ul>|
|[Settings](../reference/shared/settings/settings-object.md)|Represents custom settings for a task pane or content add-in that are stored in the host document as name/value pairs.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Access</p></li><li><p>Excel</p></li><li><p>PowerPoint</p></li><li><p>Word</p></li></ul>|
|[Slice](../reference/shared/slice/slice-object.md)|Represents a slice of a document file.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>PowerPoint</p></li><li><p>Word</p></li></ul>|
|[TableBinding](../reference/shared/binding-object/tablebinding-object/tablebinding-object.md)|Represents a binding in two dimensions of rows and columns, optionally with headers.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Access</p></li><li><p>Excel</p></li><li><p>Word</p></li></ul>|
|[TableData](../reference/shared/tabledata/tabledata-object.md)|Represents the data in a table or a  **TableBinding**.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Access</p></li><li><p>Excel</p></li><li><p>Word</p></li></ul>|
|[TextBinding](../reference/shared/binding-object/tablebinding-object/textbinding-object.md)|Represents a bound text selection in the document.|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Excel</p></li><li><p>Word</p></li></ul>|

## Supported host applications


|||
|:-----|:-----|
|**Supported hosts**|
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Access</p></li><li><p>Excel</p></li><li><p>Outlook</p></li><li><p>PowerPoint</p></li><li><p>Project</p></li><li><p>Word</p></li></ul>See "Supported host applications" in the Objects table for details about support for each object.|
|**Library**|Office.js|
|**Namespace**|Office|

## Additional resources
<a name="bk_addresources"> </a>


- [JavaScript API for Office](../reference/javascript-api-for-office.md)
    
