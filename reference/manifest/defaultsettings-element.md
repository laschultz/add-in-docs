
# DefaultSettings element
Specifies the default source location and other default settings for your content or task pane add-in .

 **Add-in type:** Content, Task pane


## Syntax:


```XML
<DefaultSettings>
  ...
</DefaultSettings>
```


## Contained in:

[OfficeApp](../reference/manifest/officeapp-element.md)


## Can contain:



|**Element**|**Content**|**Mail**|**TaskPane**|
|:-----|:-----|:-----|:-----|
|[SourceLocation](../reference/manifest/override-element.md)|x||x|
|[RequestedWidth](../reference/manifest/requestedwidth-element.md)|x|||
|[RequestedHeight](../reference/manifest/requestedheight-element.md)|x|||

## Remarks

The source location and other settings in the  **DefaultSettings** element apply only to content and task pane add-ins. For mail add-ins, you specify the default locations for source files and other default settings in the[FormSettings](../reference/manifest/formsettings-element.md) element.

