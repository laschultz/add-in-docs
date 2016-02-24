
# Enumerations (JavaScript API for Office)

You can specify an enumerated value by using either its fully qualified enumeration name ( `Office.CoercionType.Text`) or its corresponding text value ( `"text"`). For example, the following method call uses enumeration names:


```
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, {valueFormat:Office.ValueFormat.Unformatted, filterType:Office.FilterType.All}, 
   function (result) {
      if (result.status === Office.AsyncResultStatus.Success)      
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {            
         var err = result.error; 
         write(err.name + ": " + err.message);
      }
   });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


Here's the same call using the enumeration text values:




```
Office.context.document.getSelectedDataAsync("text", {valueFormat:"unformatted", filterType:"all"}, 
   function (result) {
      if (result.status === "success")      
         var dataValue = result.value; // Get selected data.
         write('Selected data is ' + dataValue);
      else {            
         var err = result.error; 
         write(err.name + ": " + err.message);
      }
   });
```


## Reference



|**Name**|**Definition**|
|:-----|:-----|
|[ActiveView](../reference/enumerations/activeview-enumeration.md)|Specifies the state of the active view of the document, for example, whether the user can edit the document.|
|[AsyncResultStatus](../reference/enumerations/asyncresultstatus-enumeration.md)|Specifies the result of an asynchronous call.|
|[AttachmentType](http://msdn.microsoft.com/library/83883a47-a937-4afb-a55e-e789057335c4%28Office.15%29.aspx)|Specifies the type of an attachment to an email message or meeting request. Outlook 2013 does not support this enumeration.|
|[BindingType](../reference/enumerations/bindingtype-enumeration.md)|Specifies the type of the binding object that should be returned.|
|[BodyType](http://msdn.microsoft.com/library/31350fe6-4c42-4cbb-a5b2-4fb2d360fa11%28Office.15%29.aspx)|Specifies the text type for the body of an appointment or message.|
|[CoercionType](../reference/enumerations/coerciontype-enumeration.md)|Specifies how to coerce data returned or set by the invoked method.|
|[CustomXMLNodeType](../reference/enumerations/customxmlnodetype-enumeration.md)|Specifies the node type.|
|[DocumentMode](../reference/enumerations/documentmode-enumeration.md)|Specifies whether the document in associated application is read-only or read-write. |
|[EntityType](http://msdn.microsoft.com/library/0035be38-8a65-4693-bcc4-0a8dd7b1495b%28Office.15%29.aspx)|Specifies an entity's type.|
|[EventType](../reference/enumerations/eventtype-enumeration.md)|Specifies the kind of event that was raised.|
|[FileType](../reference/enumerations/filetype-enumeration.md)|Specifies the format in which to return the document.|
|[GoToType](../reference/enumerations/gototype-enumeration.md)|Specifies the type of place or object to navigate to.|
|[FilterType](../reference/enumerations/filtertype-enumeration.md)|Specifies whether filtering from the host application is applied when the data is retrieved.|
|[InitializationReason](../reference/enumerations/initializationreason-enumeration.md)|Specifies whether the add-in was just inserted or was already contained in the document.|
|[ItemType](http://msdn.microsoft.com/library/e0bb23fd-f360-4b0f-b72c-1cf08d4cab3f%28Office.15%29.aspx)|Specifies an item's type.|
|[notificationMessageType](http://msdn.microsoft.com/library/ff00c89d-0019-4545-a95b-7ed0db712ce9%28Office.15%29.aspx)|Specifies the notification message for an appointment or message.|
|[ProjectProjectFields](../reference/enumerations/projectprojectfields-enumeration.md)|Specifies the project fields that are available as a parameter for the [getProjectFieldAsync](../reference/shared/projectdocument/getprojectfieldasync-method.md) method.|
|[ProjectResourceFields](../reference/enumerations/projectresourcefields-enumeration.md)|Specifies the resource fields that are available as a parameter for the [getResourceFieldAsync](../reference/shared/projectdocument/gettaskfieldasync-method.md) method.|
|[ProjectTaskFields](../reference/enumerations/projecttaskfields-enumeration.md)|Specifies the task fields that are available as a parameter for the [getTaskFieldAsync](../reference/shared/projectdocument/gettaskfieldasync-method.md) method.|
|[ProjectViewTypes](../reference/enumerations/projectviewtypes-enumeration.md)|Specifies the types of views that the [getSelectedViewAsync](../reference/shared/projectdocument/getselectedviewasync-method.md) method can recognize.|
|[RecipientType](http://msdn.microsoft.com/library/6e7c4029-6e52-47f6-98d2-4cd3ce7bd8b4%28Office.15%29.aspx)|Specifies the type of recipient for an appointment.|
|[ResponseType](http://msdn.microsoft.com/library/b3e723ca-4be0-4846-ad97-0eecab4355eb%28Office.15%29.aspx)|Specifies the response to a meeting invitation.|
|[SelectionMode](../reference/enumerations/selectionmode-enumeration.md)|Specifies whether to select (highlight) the location to navigate to (when using the [Document.goToByIdAsync](../reference/shared/document/gotobyidasync-method.md) method).|
|[SourceProperty](http://msdn.microsoft.com/library/6a209a7f-57cd-4dc3-869e-07b0f5928b28%28Office.15%29.aspx)|Specifies the source of the data returned by the invoked method.|
|[Table](../reference/enumerations/table-enumeration.md)|Specifies enumerated values for the  `cells:` property in the _cellFormat_ parameter of[table formatting methods](http://msdn.microsoft.com/library/46b05707-b350-41be-b6b8-311799c71a33%28Office.15%29.aspx).|
|[ValueFormat](../reference/enumerations/valueformat-enumeration.md)|Specifies whether values, such as numbers and dates, returned by the invoked method are returned with their formatting applied.|

## Support details
<a name="bk_support"> </a>

Support for each enumeration differs across Office host applications. See the "Support details" section of each enumerations's topic for host support information.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


|||
|:-----|:-----|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|
