
# ProjectDocument.ViewSelectionChanged event (JavaScript API for Office)
Occurs when the active view changes in the active project.

|||
|:-----|:-----|
|**Hosts:**|Project|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Selection|
|**[Added](#bk_history) in**|1.0|
[See all support details](#bk_support)

```
Office.EventType.ViewSelectionChanged
```


## Remarks

 **ViewSelectionChanged** is an[EventType](../reference/enumerations/eventtype-enumeration.md) enumeration constant that can be used in the[ProjectDocument.addHandlerAsync](../reference/shared/projectdocument/addhandlerasync-method.md) and[ProjectDocument.removeHandlerAsync](../reference/shared/projectdocument/removehandlerasync-method.md) methods to add or remove a handler for the event.


## Example

The following code example adds a handler for the  **ViewSelectionChanged** event. When the active view changes, it gets the name and type of the active view.

The example assumes your add-in has a reference to the jQuery library and that the following page control is defined in the content div in the page body.




```HTML
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
                Office.EventType.ViewSelectionChanged,
                getActiveView);
            getActiveView();
        });
    };

    // Get the name and type of the active view and display it in the add-in.
    function getActiveView() {
        Office.context.document.getSelectedViewAsync(
            function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    onError(result.error);
                }
                else {
                    var output = String.format(
                        'View name: {0}<br/>View type: {1}',
                        result.value.viewName, result.value.viewType);
                    $('#message').html(output);
                }
            }
        );
    }

    function onError(error) {
        $('#message').html(error.name + ' ' + error.code + ': ' + error.message);
    }
})();

```

For an example that shows how to use a  **ViewSelectionChanged** event handler in a Project add-in, see[Create your first task pane add-in for Project 2013 by using a text editor](http://msdn.microsoft.com/library/f6ab544a-a841-4f1b-b0c4-5001b33bba01%28Office.15%29.aspx).


## Support details
<a name="bk_support"> </a>

A capital Y in the following matrix indicates that this event is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this event.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](http://msdn.microsoft.com/library/67340567-bb9a-498c-96d3-3f52f28c16bc%28Office.15%29.aspx).


|
|
||**Office for Windows desktop**|**Office Online(in browser)**|
|:-----|:-----|:-----|
|**Project**|Y||

|||
|:-----|:-----|
|**Available in requirement sets**||
|**Minimum permission level**|[Restricted](http://msdn.microsoft.com/library/da2efadc-4ebf-45fe-be39-397ac1eb1dbd%28Office.15%29.aspx)|
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


[Create your first task pane add-in for Project 2013 by using a text editor](http://msdn.microsoft.com/library/f6ab544a-a841-4f1b-b0c4-5001b33bba01%28Office.15%29.aspx)
[EventType enumeration](../reference/enumerations/eventtype-enumeration.md)
[ProjectDocument.addHandlerAsync method](../reference/shared/projectdocument/addhandlerasync-method.md)
[ProjectDocument.removeHandlerAsync method](../reference/shared/projectdocument/removehandlerasync-method.md)
[ProjectDocument object](../reference/shared/projectdocument/projectdocument-object.md)
