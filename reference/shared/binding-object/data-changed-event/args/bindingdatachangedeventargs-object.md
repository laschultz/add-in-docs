
# BindingDataChangedEventArgs object (JavaScript API for Office)
Provides information about the binding that raised the [DataChanged](../reference/shared/binding-object/data-changed-event/bindingdatachanged-event.md) event.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**[Last changed](#bk_history) in BindingEvents**|1.1|
[See all support details](#bk_support)

```
Office.EventType.BindingDataChanged
```


## Members


**Properties**


|**Name**|**Description**|
|:-----|:-----|
|[binding](../reference/shared/binding-object/data-changed-event/args/binding-property.md)|Gets a [Binding](../reference/shared/binding-object/binding-object.md) object that represents the binding that raised the **DataChanged** event.|
|[type](../reference/shared/binding-object/data-changed-event/args/type-property.md)|Gets an [EventType](../reference/enumerations/eventtype-enumeration.md) enumeration value that identifies the kind of event that was raised.|

## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online(in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
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
|1.1|Added support for Excel and Word in Office for iPad.|
|1.1|Added support for this event in add-ins for Access.|
|1.0|Introduced|
