
# AsyncResult object (JavaScript API for Office)
An object which encapsulates the result of an asynchronous request, including status and error information if the request failed.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**[Last changed](#bk_history) in**|1.1|
[See all support details](#bk_support)

```
AsyncResult
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|**[asyncContext](../reference/shared/asyncresult/asynccontext-property.md)**|Gets the user-defined item passed to the optional  _asyncContext_ parameter of the invoked method in the same state as it was passed in.|
|**[error](../reference/shared/asyncresult/error-property.md)**|Gets an  **Error** object that provides a description of the error, if any error occurred.|
|**[status](../reference/shared/asyncresult/status-property.md)**|Gets the status of the asynchronous operation.|
|**[value](../reference/shared/asyncresult/value-property.md)**|Gets the payload or content of this asynchronous operation, if any.|

## Remarks

When the function you pass to the  _callback_ parameter of an "Async" method executes, it receives an[AsyncResult](../reference/shared/asyncresult-object.md) object that you can access from the callback function's only parameter.

The following is an example applicable to content and task pane add-ins. The example shows a call to the [getSelectedDataAsync](../reference/shared/document/getselecteddataasync-method.md) method of the **Document** object.




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
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

The anonymous function passed as the  _callback_ argument ( `function (result){...}`) has a single parameter named  _result_ that provides access to an **AsyncResult** object when the function executes. When the call to the **getSelectedDataAsync** method completes, the callback function executes, and the following line of code accesses the **value** property of the **AsyncResult** object to return the data selected in the document:

 `var dataValue = result.value;`

Note that other lines of code in the function use the  _result_ parameter of the callback function to access the **status** and **error** properties of the **AsyncResult** object.

The  **AsyncResult** object is available from the function passed as the argument to the _callback_ parameter of the following methods:



|**Parent Object**|**Method**|
|:-----|:-----|
|**Document**(Excel, PowerPoint, Project, and Word only)|[getSelectedDataAsync](../reference/shared/document/getselecteddataasync-method.md)|
||[setSelectedDataAsync](../reference/shared/document/setselecteddataasync-method.md)|
|**Bindings** (Excel and Word only)|[addFromPromptAsync](../reference/shared/bindings-object/addfrompromptasync-method.md)|
||[addFromSelectionAsync](../reference/shared/bindings-object/addfromselectionasync-method.md)|
||[getAllAsync](../reference/shared/bindings-object/getallasync-method.md)|
||[getByIdAsync](../reference/shared/bindings-object/getbyidasync-method.md)|
||[releaseByIdAsync](../reference/shared/bindings-object/releasebyidasync-method.md)|
|**Binding** (Excel and Word only)|[getDataAsync](../reference/shared/binding-object/getdataasync-method.md)|
||[setDataAsync](../reference/shared/binding-object/setdataasync-method.md)|
||[removeHandlerAsync](../reference/shared/binding-object/removehandlerasync-method.md)|
|**TableBinding** (Excel and Word only)||
||[addRowsAsync](../reference/shared/binding-object/tablebinding-object/addrowsasync-method.md)|
||[deleteAllDataValuesAsync](../reference/shared/binding-object/tablebinding-object/deletealldatavaluesasync-method.md)|
|**Settings** (Excel, PowerPoint, and Word only)|[refreshAsync](../reference/shared/settings/refreshasync-method.md)|
||[saveAsync](../reference/shared/settings/saveasync-method.md)|
|**CustomXmlNode** (Word only)|[getNodesAsync](../reference/shared/customxmlnode-object/getnodesasync-method.md)|
||[getNodeValueAsync](../reference/shared/customxmlnode-object/getnodevalueasync-method.md)|
||[getXmlAsync](../reference/shared/customxmlnode-object/getxmlasync-method.md)|
||[setNodeValueAsync](../reference/shared/customxmlnode-object/setnodevalueasync-method.md)|
||[setXmlAsync](../reference/shared/customxmlnode-object/setxmlasync-method.md)|
|**CustomXmlPart** (Word only)|[deleteAsync](../reference/shared/customxmlpart-object/deleteasync-method.md)|
||[getNodesAsync](../reference/shared/customxmlpart-object/getnodesasync-method.md)|
||[getXmlAsync](../reference/shared/customxmlpart-object/getxmlasync-method.md)|
|**CustomXmlParts** (Word only)|[addAsync](../reference/shared/customxmlparts-object/addasync-method.md)|
||[getByIdAsync](../reference/shared/customxmlparts-object/getbyidasync-method.md)|
||[getByNamespaceAsync](../reference/shared/customxmlparts-object/getbynamespaceasync-method.md)|
|**CustomXmlPrefixMappings** (Word only)|[addNamespaceAsync](../reference/shared/customxmlprefixmappings-object/addnamespaceasync-method.md)|
||[getNamespaceAsync](../reference/shared/customxmlprefixmappings-object/getnamespaceasync-method.md)|
||[getPrefixAsync](../reference/shared/customxmlprefixmappings-object/getprefixasync-method.md)|
|**Mailbox** (Outlook only)|[getUserIdentityTokenAsync](http://msdn.microsoft.com/library/c658518b-6867-41a0-99cf-810303e4c539%28Office.15%29.aspx)|
||[makeEwsRequestAsync](http://msdn.microsoft.com/library/2ec380e0-4a67-4146-92a6-6a39f65dc6f2%28Office.15%29.aspx)|
|**CustomProperties** (Outlook only)|[saveAsync](http://msdn.microsoft.com/library/690d5aa9-62b5-4e5c-9548-62dfdbb5fa56%28Office.15%29.aspx)|
|**Item** (Outlook only)|[loadCustomPropertiesAsync](http://msdn.microsoft.com/library/dfbec151-8ea7-4915-b723-09ea1396a261%28Office.15%29.aspx)|
|**RoamingSettings** (Outlook only)|[saveAsync](http://msdn.microsoft.com/library/a616f71c-a447-423f-a0d2-e9d6f1ac32f8%28Office.15%29.aspx)|

## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


||
|:-----|
|**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|**OWA for Devices**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for add-ins for Access.|
|1.0|Introduced|
