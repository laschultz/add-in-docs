
# Context.officeTheme property
Provides access to the properties for Office theme colors.

 **Important:** This API currently works only in Excel, Outlook, PowerPoint, and Word in[Office 2016 Preview](https://products.office.com/en-us/office-2016-preview) on Windows desktop.


|||
|:-----|:-----|
|**Hosts:**|Excel, Outlook, PowerPoint, Word|
|**Available in [Requirement set](http://msdn.microsoft.com/library/6b6702f2-b0a5-46ab-a356-8dda897ca8ae%28Office.15%29.aspx)**|Not in a set|
|**[Added](#bk_history) in**|1.3|

[See all support details](#bk_support)


```
Office.context.officeTheme
```


## Members


**Properties**

|||
|:-----|:-----|
|Name|Description|
|[bodyBackgroundColor ](../reference/shared/context/bodybackgroundcolor-property.md)|Gets the Office theme body background color.|
|[bodyForegroundColor](../reference/shared/context/bodyforegroundcolor-property.md)|Gets the Office theme body foreground color.|
|[controlBackgroundColor](../reference/shared/context/controlbackgroundcolor-property.md)|Gets the Office theme control background color.|
|[controlForegroundColor](../reference/shared/context/controlforegroundcolor-property.md)|Gets the Office theme control foreground color.|

## Remarks

Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with  ** File** > **Office Account** > **Office Theme** UI, which is applied across all Office host applications. Using Office theme colors is appropriate for Outlook and task pane add-ins.


## Example


```
function applyOfficeTheme(){
    // Get office theme colors.
    var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
    var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
    var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
    var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

    // Apply body background color to a CSS class.
    $('.body').css('background-color', bodyBackgroundColor);
}
```


## Support details
<a name="bk_support"> </a>


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
|1.3|Introduced|
