
# Office.select method (JavaScript API for Office)
Creates a promise to return a binding based on the selector string passed in.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement sets](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**[Last changed](#bk_history) in**|1.1|
[See all support details](#bk_support)

```
Office.select(str, onError);
```


## Parameters


-  _str_Type:  **string**
    
    The selector string to parse and create a promise for.
    

-  _onError_Type:  **function**
    
    A function that is invoked when the callback returns, whose only parameter is of type  **AsyncResult**. Optional.
    

## Callback Value

When the function you passed to the  _onError_ parameter executes, it receives an[AsyncResult](../reference/shared/asyncresult-object.md) object that you can access from the callback function's only parameter. If the operation failed, use the[AsyncResult.error](../reference/shared/asyncresult/error-property.md) property to access an[Error](../reference/shared/error/error-object.md) object that provides information about the error.


## Remarks

The  **Office.select** method provides access to a[Binding](../reference/shared/binding-object/binding-object.md) object promise that attempts to return the specified binding when any of its asynchronous methods are invoked.

Supported formats: "bindings# _bindingId_", which returns a  **Binding** object for the binding with the[id](../reference/shared/binding-object/id-property.md) of `bindingId`. For more information, see [Asynchronous programming in Office Add-ins](http://msdn.microsoft.com/library/7fe6bb42-3178-4d96-85f5-af5caea7b950%28Office.15%29.aspx#AsyncProgramming_PromisesPattern) and[Bind to regions in a document or spreadsheet](http://msdn.microsoft.com/library/5bf788db-d788-4d91-bcb6-fc3913b40012%28Office.15%29.aspx).


 **Note**  If the  **select** method promise successfully returns a **Binding** object, that object exposes only the following four methods of the[Binding](../reference/shared/binding-object/binding-object.md) object:[getDataAsync](../reference/shared/binding-object/getdataasync-method.md), [setDataAsync](../reference/shared/binding-object/setdataasync-method.md), [addHandlerAsync](../reference/shared/binding-object/addhandlerasync-method.md), and [removeHandlerAsync](../reference/shared/binding-object/removehandlerasync-method.md). If the promise cannot return a  **Binding** object, the _onError_ callback can be used to access an[asyncResult.error](../reference/shared/asyncresult/error-property.md) object to get more information.If you need to call a member of the  **Binding** object other than the four methods exposed by the **Binding** object promise returned by the **select** method, instead use the[getByIdAsync](../reference/shared/bindings-object/getbyidasync-method.md) method by using the[Document.bindings](../reference/shared/document/bindings-property.md) property and[Bindings.getByIdAsync](../reference/shared/bindings-object/getbyidasync-method.md) method to retrieve the **Binding** object.


## Example

The following code example uses the  **select** method to retrieve a binding with the **id** " `cities`" from the  **Bindings** collection, and then calls the[addHandlerAsync](../reference/shared/binding-object/addhandlerasync-method.md) method to add an event handler for the[dataChanged](../reference/shared/binding-object/data-changed-event/bindingdatachanged-event.md) event of the binding.


```
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```




## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


|
|
||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Access**||Y||
|**Excel**|Y|Y|Y|
|**Word**|Y||Y|

|||
|:-----|:-----|
|**Available in requirement sets**|MatrixBindings, PartialTableBindings, TableBindings, TextBindings|
|**Minimum permission level**|[ReadDocument (ReadAllDocument for Open Office XML)](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history
<a name="bk_history"> </a>


****


|**Version**|**Changes**|
|:-----|:-----|
|1.1|Added support for Excel and Word in Office for iPad|
|1.1|Added the use of the  **select** method to return table bindings created in content add-ins for Access.|
|1.0|Introduced|
