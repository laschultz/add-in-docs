
# Labs.registerDeserializer

 _**Applies to:** apps for Office | Office Add-ins | Office Mix | PowerPoint_

Deserializes a specified JSON object into an object. Should be used by component authors only.

```
function registerDeserializer(type: string, deserialize: (json: Core.ILabObject): any): void
```


## Parameters


|||
|:-----|:-----|
|json|The [Labs.Core.ILabObject](../powerpoint/office-mix/reference/labs.core.ilabobject.md) to deserialize.|

## Return value

Returns an [Labs.Core.ILabObject](../powerpoint/office-mix/reference/labs.core.ilabobject.md) instance.

