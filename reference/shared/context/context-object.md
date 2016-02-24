
# Context object (JavaScript API for Office)
Represents the runtime environment of the add-in and provides access to key objects of the API.

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Outlook, PowerPoint, Project, Word|
|**[Last changed](#bk_history) in**|1.1|
[See all support details](#bk_support)

```
Office.context
```


## Members


**Properties**

|||
|:-----|:-----|
|Name|Description|
|[commerceAllowed](../reference/shared/context/commerceallowed-property.md)|Gets whether the add-in is running on a platform that allows links to external payment systems.|
|[contentLanguage](../reference/shared/context/contentlanguage-property.md)|Gets the locale (language) for data as it is stored in the document or item.|
|[displayLanguage](../reference/shared/context/displaylanguage-property.md)|Gets the locale (language) for the UI of the hosting application.|
|[document](../reference/shared/context/document-property.md)|Gets an object that represents the document the content or task pane add-in is interacting with.|
|[mailbox](../reference/shared/context/mailbox-property.md)|Gets the  **mailbox** object that provides access to members of the API that are specifically for Outlook add-ins.|
|[officeTheme](../reference/shared/context/officetheme-property.md)|Provides access to the properties for Office theme colors|
|[roamingSettings](../reference/shared/context/roamingsettings-property.md)|Gets an object that represents the saved custom settings of the add-in.|
|[touchEnabled](../reference/shared/context/touchenabled-property.md)|Gets whether the add-in is running in an Office host application that is touch enabled.|

## Remarks

The  **Context** object provides access to key objects in the JavaScript API for Office.


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
|1.1|Added  **commerceAllowed** and **touchEnabledAdded** properties (Excel, PowerPoint and Word on Office for iPad only).|
|1.1|Added support for add-ins with Excel and Word on Office for iPad.|
|1.1|For [contentLanguage](../reference/shared/context/contentlanguage-property.md), [displayLanguage](../reference/shared/context/displaylanguage-property.md), and [document](../reference/shared/context/document-property.md), added support for content add-ins for Access.|
|1.0|Introduced|
