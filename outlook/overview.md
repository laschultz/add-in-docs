
# Overview of Outlook add-ins architecture and features
Learn about Outlook add-ins architecture and features.

 _**Applies to:** apps for Office | Office Add-ins | Outlook_


## Architecture

An Outlook add-in consists of an XML manifest and code (JavaScript and HTML). The manifest specifies the name and description of the add-in, as well as how the add-in integrates into Outlook. Using the manifest, developers can place buttons on command surfaces, link off regular expression matches and so on. The manifest also defines the URL that hosts the JavaScript and HTML code for the add-in.

When a user or administrator acquires an add-in, the add-in's manifest is saved to the user's mailbox or into the organization. When Outlook starts up, it loads all manifests the user has installed, processes them and sets up all the extension points for the add-in (for example, display buttons in command surfaces, run regular expression against currently selected message, and so forth). The user can now use the add-in.

When the user interacts with the add-in, the JavaScript and HTML files are loaded from the host location specified in the manifest.

Add-ins use the Office.js API to access the Outlook Add-in API and interact with Outlook.


**Interaction of typical components when the user starts Outlook**

![Flow of events when starting Outlook mail app](../images/olowawecon15_LoadingDOMAgaveRuntime.png)
### Versioning

As we evolve Outlook clients and the add-in platform and add new ways for add-ins to integrate, sometimes we're unable to implement a feature at the same time across all clients (Mac, Windows, web, mobile). To handle this situation, we version both the manifest and the APIs. In this way the platform supports backwards compatibility at all times, meaning that developers can build an add-in which works in a down-level way in older clients, but also lets you take advantage of new features in newer clients. You can read more about how versioning works in [Outlook add-in manifests](../outlook/manifests/manifests.md).


## Outlook add-in features

Outlook add-ins offer many rich features that can be used to support various scenarios.


|
|
|**Feature**|**Description**|
|:-----|:-----|
|Contextual activation|Outlook add-ins are contextually based. They can activate based on the following criteria: 
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>(default) for any item in the mailbox or calendar</p></li><li><p>for a specific item type (an email message, meeting request message, or appointment)</p></li><li><p>for an item message class</p></li><li><p>for specific entities in a message or appointment, see <span sdata="link"><a href="2cd5d8f1-69b3-4a2a-b31e-81a07a7cdd9f.htm">Contextual Outlook add-ins</a></span>. </p></li><li><p>based on specific rules or regular expressions, see <span sdata="link"><a href="b3fd6d69-b968-461d-a40e-6063f4febfe6.htm">Activation rules for Outlook add-ins</a></span> and <span sdata="link"><a href="93504f92-896f-4c80-9205-ba0b125f4290.htm">Use regular expression activation rules to show an Outlook add-in</a></span>. </p></li><li><p>for string matches of properties, see <span sdata="link"><a href="a6b0904b-afe9-4882-9136-3d8cfd57fcf8.htm">Match strings in an Outlook item as well-known entities</a></span></p></li></ul>|
|Add-in commands|Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon. They are only available for add-ins that apply to all emails or events. For more information see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md). |
|Roaming settings|An Outlook add-in can save data that is specific to the user's mailbox that you can access in a subsequent Outlook session. For more information see [Get and set add-in metadata for an Outlook add-in](../outlook/apis/metadata-for-an-outlook-add-in.md). |
|Custom properties|An Outlook add-in can save data that is specific to an item in the user's mailbox that you can access in a subsequent Outlook session. For more information see [Get and set add-in metadata for an Outlook add-in](../outlook/apis/metadata-for-an-outlook-add-in.md).|
|Getting attachments or the entire selected item|An Outlook add-in can access attachments and the entire selected item from the server-side. See the following:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p>Attachments - see <span sdata="link"><a href="0f872924-ea1a-4aa2-bb7b-e12d31014612.htm">Get attachments of an Outlook item from the server</a></span> and <span sdata="link"><a href="62669c4d-6829-4476-bac2-cac95fc0961e.htm">Add and remove attachments to an item in a compose form in Outlook</a></span></p></li><li><p>Entire selected item - this is similar to using a callback token to get attachments. See the following:</p><ul><li><p><a href="https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html(Office.15).aspx#getCallbackTokenAsync" target="_blank">mailbox.getCallbackTokenAsync</a> method - provides a callback token to identify the add-in's server side code for the Exchange Server.</p></li><li><p><a href="https://dev.outlook.com/reference/add-ins/Office.context.mailbox.item.html(Office.15).aspx#itemId" target="_blank">item.itemId</a> property - identifies the item that the user is reading and that the server-side code is getting.</p></li><li><p><a href="https://dev.outlook.com/reference/add-ins/Office.context.mailbox.html(Office.15).aspx#ewsUrl" target="_blank">mailbox.ewsUrl</a> property - provides the EWS endpoint URL which, together with the callback token and item ID, the server-side code can use to access the <a href="http://msdn.microsoft.com/en-us/library/e3590b8b-c2a7-4dad-a014-6360197b68e4(Office.15).aspx" target="_blank">GetItem</a> EWS operation to get the entire item.</p></li></ul></li></ul>|
|User profile|A mail add-in can access the display name, email address, and time zone in the user's profile. For more information see the [UserProfile](https://dev.outlook.com/reference/add-ins/Office.context.mailbox.userProfile.html%28Office.15%29.md) object.|

## Get started building Outlook add-ins

To get started building Outlook add-ins, see [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx).


## Additional resources
<a name="off15con_AgaveFundOutlookAdditionalRsc"> </a>

For concepts that are applicable to developing Office Add-ins in general, see the following:


- [Design guidelines for Office Add-ins](../design/add-in-design.md)
    
- [Best practices for developing Office Add-ins](../design/add-in-development-best-practices.md)
    
- [License your Office and SharePoint Add-ins](http://msdn.microsoft.com/library/3e0e8ff6-66d6-44ff-b0c2-59108ebd9181%28Office.15%29.aspx)
    
- [Submit Office and SharePoint Add-ins and Office 365 web apps to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx)
    
- [JavaScript API for Office](http://msdn.microsoft.com/EN-US/library/fp142185%28v=office.15%29.aspx(Office.15).aspx)
    
- [Mail add-in manifests](../outlook/manifests/manifests.md)
    
