
# Create Outlook add-ins for read forms
Build Outlook add-ins for the read scenario and find information for available API features. These features include contextual activation based on message class, the existence of an attachment, or string matching by regular expressions or well-known entities, and getting attachments of an item.

 _**Applies to:** apps for Office | Office Add-ins | Outlook_


## Setting up Outlook add-ins for read forms
<a name="mod_off15_CreatingForRead_SettingUp"> </a>

Read add-ins are Outlook add-ins that are activated in the Reading Pane or read inspector in Outlook. Unlike compose add-ins (Outlook add-ins that are activated when a user is creating a message or appointment), read add-ins are available when users:


- View an email message, meeting request, meeting response, or meeting cancellation.*
    
- View a meeting item in which the user is an attendee.
    
- View a meeting item in which the user is the organizer (RTM release of Outlook 2013 and Exchange 2013 only).
    
     **Note**  Starting in the Office 2013 SP1 release, if the user is viewing a meeting item that the user has organized, only compose add-ins can activate and be available. Read add-ins are no longer available in this scenario.
* Outlook doesn't activate add-ins in read form for certain types of messages, including items that are attachments to another message, items in the Outlook Drafts or Junk Email folder, or items that are encrypted or protected in other ways.

In each of these read scenarios, Outlook activates add-ins when their activation conditions are fulfilled, and users can choose and open activated add-ins in the add-in bar in the Reading Pane or read inspector. Figure 1 shows the  **Bing Maps** add-in activated and opened as the user is reading a message that contains a geographic address.


**Figure 1. The add-in pane showing the Bing Maps add-in in action for the selected Outlook message that contains an address**

![Bing Map mail app in Outlook](../images/off15appsdk_BingMapMailAppScreenshot.jpg)


## Types of add-ins available in read mode
<a name="mod_off15_CreatingForRead_SettingUp"> </a>

Read add-ins can be any combination of the following types.


- [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md)
    
- [Contextual Outlook add-ins](../outlook/contextual-outlook-add-ins.md)
    
- [Custom pane Outlook add-ins](../outlook/custom-pane-outlook-add-ins.md)
    

## API features available to read add-ins
<a name="mod_off15_CreatingForRead_APIFeatures"> </a>

For a list of features that the JavaScript API for Office provides to Outlook add-ins in read forms, see Tables 1 and 2 in [Mail app features per version](http://msdn.microsoft.com/library/f34e2f44-8c9d-4e90-b1d7-3f29506adb92%28Office.15%29.aspx). 

See also:


- For activating add-ins in read forms: see Table 1 in [Specify activation rules in a manifest](../outlook/manifests/activation-rules.md#MailAppDefineRules_Manifest).
    
- [Use regular expression activation rules to show an Outlook add-in](../outlook/read/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Match strings in an Outlook item as well-known entities](../outlook/read/match-strings-in-an-item-as-well-known-entities.md)
    
- [Extract entity strings from an Outlook item](../outlook/read/extract-entity-strings-from-an-item.md)
    
- [Get attachments of an Outlook item from the server](../outlook/read/get-attachments-of-an-outlook-item.md)
    

## In this section
<a name="mod_off15_CreatingForRead_InThisSection"> </a>


- [Use regular expression activation rules to show an Outlook add-in](../outlook/read/use-regular-expressions-to-show-an-outlook-add-in.md)
    
- [Match strings in an Outlook item as well-known entities](../outlook/read/match-strings-in-an-item-as-well-known-entities.md)
    
- [Extract entity strings from an Outlook item](../outlook/read/extract-entity-strings-from-an-item.md)
    
- [Get attachments of an Outlook item from the server](../outlook/read/get-attachments-of-an-outlook-item.md)
    

## Additional resources
<a name="mod_off15_CreatingForRead_AdditionalResources"> </a>


- [Get Started with Outlook add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
- [Outlook add-ins](../outlook/outlook-add-ins.md)
    