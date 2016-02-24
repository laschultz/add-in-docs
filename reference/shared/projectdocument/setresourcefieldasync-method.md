
# ProjectDocument.setResourceFieldAsync method (JavaScript API for Office v1.1)
Asynchronously sets the value of the specified field for the specified resource.
 **Important:** This API works only in Project 2016 on Windows desktop.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Selection|
|**[Added](#bk_history) in**|1.1|
[See all support details](#bk_support)

```
Office.context.document.setResourceFieldAsync(resourceId, fieldId, fieldValue[, options][, callback]);
```


## Parameters


-  _resourceId_The GUID of the resource. Required.
    
-  _fieldId_The ID of the target field, as a [ProjectResourceFields](../reference/enumerations/projectresourcefields-enumeration.md) constant or its corresponding integer value. Required.
    
-  _fieldValue_The value for the target field, as  **string**,  **number**,  **boolean**, or  **object**. Required.
    
-  _options_The following [optional parameter](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters):
    
||
|:-----|
|<dl class="authored" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><dt><span class="parameter" sdata="paramReference">asyncContext</span></dt><dd><p>Type: <span class="keyword">array</span>, <span class="keyword">boolean</span>, <span class="keyword">null</span>, <span class="keyword">number</span>, <b>object</b> , <span class="keyword">string</span>, or <span class="keyword">undefined</span></p><p>A user-defined item of any type that is returned in the <a href="540c114f-0398-425c-baf3-7363f2f6bc47.htm">AsyncResult</a> object without being altered. Optional.</p><p>For example, you can pass the <span class="parameter" sdata="paramReference">asyncContext</span> argument by using the format <span class="code">{asyncContext: 'Some text'}</span> or <span class="code">{asyncContext: <object>}</span>.</p></dd></dl>|
-  _callback_Type:  **function**
    
    A function that is invoked when the method call returns, where the only parameter is of type [AsyncResult](../reference/shared/asyncresult-object.md). Optional.
    

## Callback Value

When the  _callback_ function executes, it receives an[AsyncResult](../reference/shared/asyncresult-object.md) object that you can access from the parameter in the callback function.

For the  **setResourceFieldAsync** method, the returned[AsyncResult](../reference/shared/asyncresult-object.md) object contains following properties.


****


|**Name**|**Description**|
|:-----|:-----|
|[asyncContext](../reference/shared/asyncresult/asynccontext-property.md)|The data passed in the optional  _asyncContext_ parameter, if the parameter was used.|
|[error](../reference/shared/asyncresult/error-property.md)|Information about the error, if the  **status** property equals **failed**.|
|[status](../reference/shared/asyncresult/status-property.md)|The  **succeeded** or **failed** status of the asynchronous call.|
|[value](../reference/shared/asyncresult/value-property.md)|This method does not return a value.|

## Remarks

First call the [getSelectedResourceAsync](../reference/shared/projectdocument/getselectedtaskasync-method.md) or[getResourceByIndexAsync](../reference/shared/projectdocument/getresourcebyindexasync-method.md) method to get the resource GUID, and then pass the GUID as the _resourceId_ argument to **setResourceFieldAsync**. Only a single field for a single resource can be updated in each asynchronous call.


## Example

The following code example calls [getSelectedResourceAsync](../reference/shared/projectdocument/getselectedtaskasync-method.md) to get the GUID of the resource that's currently selected in a resource view. Then it sets two resource field values by calling **setResourceFieldAsync** recursively.

The [getSelectedTaskAsync](../reference/shared/projectdocument/getselectedtaskasync-method.md) method used in the example requires that a task view (for example, Task Usage) is the active view and that a task is selected. See the[addHandlerAsync](../reference/shared/projectdocument/addhandlerasync-method.md) method for an example that activates a button based on the active view type.

The example assumes your add-in has a reference to the jQuery library and that the following page controls are defined in the content div in the page body.




```HTML
<input id="set-info" type="button" value="Set info" /><br />
<span id="message"></span>
```




```
(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {

            // After the DOM is loaded, add-in-specific code can run.
            app.initialize();
            $('#set-info').click(setResourceInfo);
        });
    };

    // Get the GUID of the resource, and then get the resource fields.
    function setResourceInfo() {
        getResourceGuid().then(
            function (data) {
                setResourceFields(data);
            }
        );
    }

    // Get the GUID of the selected resource.
    function getResourceGuid() {
        var defer = $.Deferred();
        Office.context.document.getSelectedResourceAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    defer.resolve(result.value);
                }
            }
        );
        return defer.promise();
    }

    // Set the specified fields for the selected resource.
    function setResourceFields(resourceGuid) {
        var targetFields = [Office.ProjectResourceFields.StandardRate, Office.ProjectResourceFields.Notes];
        var fieldValues = [.28, 'Notes for the resource.'];

        // Set the field value. If the call is successful, set the next field.
        for (var i = 0; i < targetFields.length; i++) {
            Office.context.document.setResourceFieldAsync(
                resourceGuid,
                targetFields[i],
                fieldValues[i],
                function (result) {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        i++;
                    }
                    else {
                        onError(result.error);
                    }
                }
            );
        }
        $('#message').html('Field values set');
    }

    function onError(error) {
        app.showNotification(error.name + ' ' + error.code + ': ' + error.message);
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
|**Available in requirement sets**||
|**Minimum permission level**|[WriteDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Introduced|

## See also
<a name="bk_history"> </a>


#### Other resources


[getSelectedResourceAsync](../reference/shared/projectdocument/getselectedtaskasync-method.md)
[getResourceByIndexAsync](../reference/shared/projectdocument/getresourcebyindexasync-method.md)
[AsyncResult object](../reference/shared/asyncresult-object.md)
[ProjectResourceFields enumeration](../reference/enumerations/projectresourcefields-enumeration.md)
[ProjectDocument object](../reference/shared/projectdocument/projectdocument-object.md)
