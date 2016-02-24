
# Error object (JavaScript API for Office)
Provides specific information about an error that occurred during an asynchronous data operation.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**[Last changed](#bk_history) in**|1.1|
[See all support details](#bk_support)

```
asyncResult.error
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[code](../reference/shared/error/code-property.md)|Gets the numeric code of the error.|
|[name](../reference/shared/error/name-property.md)|Gets the name of the error.|
|[message](../reference/shared/error/message-property.md)|Gets a detailed description of the error.|

## Remarks

The  **Error** object is accessed from the[AsyncResult](../reference/shared/asyncresult-object.md) object that is returned in the function passed as the _callback_ argument of an asynchronous data operation, such as the[setSelectedDataAsync](../reference/shared/document/setselecteddataasync-method.md) method of the **Document** object.


## Example

The following example uses the  **setSelectedDataAsync** method to set the selected text to "Hello World!", and if that fails, displays the values of the **name** and **message** properties of the **Error** object.


```
function setText() {

    Office.context.document.setSelectedDataAsync("Hello World!", {},
        function (asyncResult) {
            if (asyncResult.status === "failed")
            var err = asyncResult.error; 
                write(err.name + ": " + err.message);
        });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```




## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


|
|
||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|**OWA for Devices**|**Outlook for Mac**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|**Access**||Y||||
|**Excel**|Y|Y|Y|||
|**Outlook**|Y|Y||Y|Y|
|**PowerPoint**|Y|Y|Y|||
|**Project**|Y|||||
|**Word**|Y|Y|Y|||

|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel, PowerPoint, and Word in Office for iPad.|
|1.1|Added support for content add-ins for Access.|
|1.0|Introduced|
