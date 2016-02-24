
# Override element
Provides a way to specify the value of a setting for an additional locale.

 **Add-in type:** Content, Task pane, Mail


## Syntax:


```XML
<Override Locale="string " Value="string " />
```


## Contained in:


||
|:-----|
|[CitationText](../reference/manifest/citationtext-element.md)|
|[Description](../reference/manifest/description-element.md)|
|[DictionaryName](../reference/manifest/dictionaryname-element.md)|
|[DictionaryHomePage](../reference/manifest/dictionaryhomepage-element.md)|
|[DisplayName](../reference/manifest/displayname-element.md)|
|[HighResolutionIconUrl](../reference/manifest/highresolutioniconurl-element.md)|
|[IconUrl](../reference/manifest/iconurl-element.md)|
|[QueryUri](../reference/manifest/queryuri-element.md)|
|[SourceLocation](../reference/manifest/sourcelocation-element.md)|
|[SupportUrl](../reference/manifest/supporturl-element.md)|

## Attributes



|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|Locale|string|required|Specifies the culture name of the locale for this override in the BCP 47 language tag format, such as  `"en-US"`.|
|Value|string|required|Specifies value of the setting expressed for the specified locale.|

## Additional resources
<a name="MailAppDefineRules_AdditionalResources"> </a>


- [Localization for Office Add-ins](http://msdn.microsoft.com/library/5a1a1cd7-b716-4597-b51f-fa70357d0833%28Office.15%29.aspx#off15wecon_LocalesManifest)
    
