
# Get and set item data in a compose form in Outlook
Learn how to get or set various properties of an item in an Outlook add-in in a compose scenario, including its recipients, subject, body, and appointment location and time.

 _**Applies to:** apps for Office | Office Add-ins | Outlook_


## Getting and setting item properties for a compose add-in
<a name="mod_off15_GettingSettingItemDataCompose_GettingSettingItemProps"> </a>

In a compose form, you can get most of the properties that are exposed on the same kind of item as in a read form (such as attendees, recipients, subject, and body), and you can get a few extra properties that are relevant in only a compose form but not a read form (body, bcc). 

For most of these properties, because it's possible that an Outlook add-in and the user can be modifying the same property in the user interface at the same time, the methods to get and set them are asynchronous. Table 1 lists the item-level properties and corresponding asynchronous methods to get and set them in a compose form. The [item.itemType](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) and[item.conversationId](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md) properties are exceptions because users cannot modify them. You can programmatically get them the same way in a compose form as in a read form, directly from the parent object.

Other than accessing item properties in the JavaScript API for Office, you can access item-level properties using Exchange Web Services (EWS). With the  **ReadWriteMailbox** permission, you can use the[mailbox.makeEwsRequestAsync](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.html%28Office.15%29.md) method to access EWS operations,[GetItem](http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4%28Office.15%29.aspx) and[UpdateItem](http://msdn.microsoft.com/en-us/library/5d027523-e0bc-4da2-b60b-0cb9fc1fdfe4%28Office.15%29.aspx), to get and set more properties of an item or items in the user's mailbox.  **makeEwsRequestAsync** is available in both compose and read forms. For more information about the **ReadWriteMailbox** permission, and accessing EWS through the Office Add-ins platform, see[Understanding Outlook add-in permissions](../outlook/privacy/understanding-outlook-add-in-permissions.md) and[Call web services from an Outlook add-in](../outlook/web-services.md).


**Table 1. Asynchronous methods to get or set item properties in a compose form**


|**Property**|**Property type**|**Asynchronous method to get**|**Asynchronous method(s) to set**|
|:-----|:-----|:-----|:-----|
|[bcc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Recipients](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|[Recipients.getAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|[Recipients.addAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)[Recipients.setAsync](http://dev.outlook.com/reference/add-ins/Recipients.html%28Office.15%29.md)|
|[body](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Body](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|[Body.getAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|[Body.prependAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)[Body.setAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)[Body.setSelectedDataAsync](http://dev.outlook.com/reference/add-ins/Body.html%28Office.15%29.md)|
|[cc](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Recipients**|**Recipients.getAsync**|**Recipients.addAsync** **Recipients.setAsync**|
|[end](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Time](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|[Time.getAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|[Time.setAsync](http://dev.outlook.com/reference/add-ins/Time.html%28Office.15%29.md)|
|[location](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Location](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|[Location.getAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|[Location.setAsync](http://dev.outlook.com/reference/add-ins/Location.html%28Office.15%29.md)|
|[optionalAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Recipients**|**Recipients.getAsync**|**Recipients.addAsync** **Recipients.setAsync**|
|[requiredAttendees](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Recipients**|**Recipients.getAsync**|**Recipients.addAsync** **Recipients.setAsync**|
|[start](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Time**|**Time.getAsync**|**Time.setAsync**|
|[subject](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|[Subject](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|[Subject.getAsync](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|[Subject.setAsync](http://dev.outlook.com/reference/add-ins/Subject.html%28Office.15%29.md)|
|[to](http://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html%28Office.15%29.md)|**Recipients**|**Recipients.getAsync**|**Recipients.addAsync** **Recipients.setAsync**|

## In this section
<a name="mod_off15_GettingSettingItemDataCompose_InThisSection"> </a>

See the following examples that show how to asynchronously get or set some of the item properties.


- [Get, set, or add recipients when composing an appointment or message in Outlook](../outlook/compose/get-set-or-add-recipients.md)
    
- [Get or set the subject when composing an appointment or message in Outlook](../outlook/compose/get-or-set-the-subject.md)
    
- [Insert data in the body when composing an appointment or message in Outlook](../outlook/compose/insert-data-in-the-body.md)
    
- [Get or set the location when composing an appointment in Outlook](../outlook/compose/get-or-set-the-location-of-an-appointment.md)
    
- [Get or set the time when composing an appointment in Outlook](../outlook/compose/get-or-set-the-time-of-an-appointment.md)
    

## Additional resources
<a name="mod_off15_GettingSettingItemDataCompose_AdditionalRsc"> </a>


- [Create Outlook add-ins for compose forms](../outlook/compose/compose-scenario.md)
    
- [Understanding Outlook add-in permissions](../outlook/privacy/understanding-outlook-add-in-permissions.md)
    
- [Call web services from an Outlook add-in](../outlook/web-services.md)
    
- [Get and set Outlook item data in read or compose forms](../outlook/apis/item-data.md)
    


