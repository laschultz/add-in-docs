
# ProjectDocument.removeHandlerAsync method (JavaScript API for Office)
Asynchronously removes an event handler for the task selection changed event in a [ProjectDocument](../reference/shared/projectdocument/projectdocument-object.md) object.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Selection|
|**[Added](#bk_history) in**|1.0|
[See all support details](#bk_support)

```
Office.context.document.removeHandlerAsync(eventType[, options][, callback]);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _eventType_|[EventType](../reference/enumerations/eventtype-enumeration.md)|The type of event to remove, as an [EventType](../reference/enumerations/eventtype-enumeration.md) constant or its corresponding text value. Required.The following table shows valid  _eventType_ arguments for a[ProjectDocument](../reference/shared/projectdocument/projectdocument-object.md) object.

|**Enumeration**|**Text value**|
|:-----|:-----|
|[Office.EventType.ResourceSelectionChanged](../reference/shared/projectdocument/resourceselectionchanged-event.md)|resourceSelectionChanged|
|[Office.EventType.TaskSelectionChanged](../reference/shared/projectdocument/taskselectionchanged-event.md)|taskSelectionChanged|
|[Office.EventType.ViewSelectionChanged](../reference/shared/projectdocument/viewselectionchanged-event.md)|viewSelectionChanged|
||
| _options_|**object**|Specifies any of the following [optional parameters](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters).||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the  _callback_ function executes, it receives an[AsyncResult](../reference/shared/asyncresult-object.md) object that you can access from the parameter in the callback function.

For the  **removeHandlerAsync** method, the returned[AsyncResult](../reference/shared/asyncresult-object.md) object contains the following properties.


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../reference/shared/asyncresult/asynccontext-property.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../reference/shared/asyncresult/error-property.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../reference/shared/asyncresult/status-property.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../reference/shared/asyncresult/value-property.md)|**removeHandlerAsync** always returns **undefined**.|

## Example

The following code example uses [addHandlerAsync](../reference/shared/projectdocument/addhandlerasync-method.md) to add an event handler for the[ResourceSelectionChanged](../reference/shared/projectdocument/resourceselectionchanged-event.md) event and **removeHandlerAsync** to remove the handler.

When a resource is selected in a resource view, the handler displays the resource GUID. When the handler is removed, the GUID is not displayed.

The example assumes your add-in has a reference to the jQuery library and that the following page control is defined in the content div in the page body.




```HTML
<input id="remove-handler" type="button" value="Remove handler" /><br />
<span id="message"></span>
```




```
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            Office.context.document.addHandlerAsync(
                Office.EventType.ResourceSelectionChanged,
                getResourceGuid);
            $('#remove-handler').click(removeEventHandler);
        });
    };

    // Remove the event handler.
    function removeEventHandler() {
        Office.context.document.removeHandlerAsync(
            Office.EventType.ResourceSelectionChanged,
            {handler:getResourceGuid,
            asyncContext:'The handler is removed.'},
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#remove-handler').attr('disabled', 'disabled');
                    $('#message').html(result.asyncContext);
                }
            }
        );
    }

    // Get the GUID of the currently selected resource and display it in the add-in.
    function getResourceGuid() {
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    $('#message').html('Resource GUID: ' + result.value);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();
```


## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


|
|
||**Office for Windows desktop**|**Office Online(in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Available in requirement sets**|Selection|
|**Minimum permission level**|[ReadWriteDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|

## See also
<a name="bk_history"> </a>


#### Other resources


[addHandlerAsync method](../reference/shared/projectdocument/addhandlerasync-method.md)
[EventType enumeration](../reference/enumerations/eventtype-enumeration.md)
[ProjectDocument object](../reference/shared/projectdocument/projectdocument-object.md)
