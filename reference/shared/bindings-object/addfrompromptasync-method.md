
# Bindings.addFromPromptAsync method (JavaScript API for Office)
 Displays UI that lets the user specify a selection to bind to.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Not in a set|
|**[Last changed](#bk_history)**|1.1|
[See all support details](#bk_support)

[![Try out this call in the interactive API Tutorial for Excel](../images/819b84bf-151c-4a12-80c3-d6f8d7c03251.png)](http://officeapitutorial.azurewebsites.net/Redirect.mdl?scenario=Get+Selected+Coordinates)


```
_bindingsObj.addFromPromptAsync(bindingType [, options], callback);
```


## Parameters



|**Name**|**Type**|**Description**|**Support notes**|
|:-----|:-----|:-----|:-----|
| _bindingType_|[BindingType](../reference/enumerations/bindingtype-enumeration.md)|Specifies the type of the binding object to create. Required. Returns  **null** if the selected object cannot be coerced into the specified type.||
| _options_|**object**|Specifies any of the following [optional parameters](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_OptionalParameters)||
| _id_|**string**|Specifies the unique name to be used to identify the new binding object.If no argument is passed for the  _id_ parameter, the[Binding.id](../reference/shared/binding-object/id-property.md) is autogenerated.||
| _promptText_|**string**|Specifies the string to display in the prompt UI that tells the user what to select. Limited to 200 characters. If no  _promptText_ argument is passed, "Please make a selection" is displayed.||
| _sampleData_|[TableData](../reference/shared/tabledata/tabledata-object.md)|Specifies a table of sample data displayed in the prompt UI as an example of the kinds of fields (columns) that can be bound by your add-in. The headers provided in the  **TableData** object specify the labels used in the field selection UI. Optional. **Note:** This parameter is used only in add-ins for Access. It is ignored if provided when calling the method in an add-in for Excel.||
| _asyncContext_|**array**,  **boolean**,  **null**,  **number**,  **object**, **string**, or  **undefined**|A user-defined item of any type that is returned in the  **AsyncResult** object without being altered.||
| _callback_|**object**|A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**.||

## Callback Value

When the function you passed to the  _callback_ parameter executes, it receives an[AsyncResult](../reference/shared/asyncresult-object.md) object that you can access from the callback function's only parameter.

In the callback function passed to the  **addFromPromptAsync** method, you can use the properties of the **AsyncResult** object to return the following information.



|**Property**|**Use to...**|
|:-----|:-----|
|[AsyncResult.value](../reference/shared/asyncresult/value-property.md)|Access the [Binding](../reference/shared/binding-object/binding-object.md) object that represents the selection specified by the user.|
|[AsyncResult.status](../reference/shared/asyncresult/status-property.md)|Determine the success or failure of the operation.|
|[AsyncResult.error](../reference/shared/asyncresult/error-property.md)|Access an [Error](../reference/shared/error/error-object.md) object that provides error information if the operation failed.|
|[AsyncResult.asyncContext](../reference/shared/asyncresult/asynccontext-property.md)|Access your user-defined  **object** or value, if you passed one as the _asyncContext_ parameter.|

## Remarks

Adds a binding object of the specified type to the [Bindings](../reference/shared/bindings-object/bindings-object.md) collection, which will be identified with the supplied _id_. The method fails if the specified selection cannot be bound.


## Example




```
function addBindingFromPrompt() {

    Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Text, { id: 'MyBinding', promptText: 'Select text to bind to.' }, function (asyncResult) {
        write('Added new binding with type: ' + asyncResult.value.type + ' and id: ' + asyncResult.value.id);
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


||
|:-----|
|**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Not in a set|
|**Minimum permission level**|[ReadDocument](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>



|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel in Office for iPad.|
|1.1|In add-ins for Excel, you can create a table binding (passing  _bindingType_ as **Office.BindingType.Table**) for a range of cells that contains tabular data even when that data was not added to the spreadsheet as a table in the Excel UI (by using the  **Insert** > **Tables** > **Table** or **Home** > **Styles** > **Format as Table** commands).|
|1.1|Added support for table binding in content add-ins for Access. |
|1.1|Added support for binding to matrix data as a table binding in add-ins for Excel.|
|1.0|Introduced|
