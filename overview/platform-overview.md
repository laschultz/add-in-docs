
# Office Add-ins platform overview
Use the Office Add-ins platform to create engaging new consumer and enterprise experiences for Office client applications. Using the power of the web and standard web technologies like HTML5, XML, CSS3, JavaScript, and REST APIs, create add-ins that interact with Office documents, email messages, meeting requests, and appointments.

 _**Applies to:** Access apps for SharePoint | apps for Office | Excel | Office Add-ins | Outlook | PowerPoint | Project | Word_

This article provides a quick overview of the Office Add-ins platform and how an add-in works with an Office application. To find out how to start developing add-ins right away, see [Development basics](#StartBuildingApps_DevelopmentBasics). 

An Office Add-in is a web application hosted in a web browser control or iframe running in the context of an Office host application that can interact with a user's documents or mail items. You can use Office Add-ins to extend and interact with: 


-  **Documents or data -** Word documents, Excel spreadsheets, PowerPoint presentations, Access browser-based databases, and Project schedules and views.
    
-  **Outlook mailbox items -** Email messages, meeting requests, or appointments.
    
Add-ins can run in multiple environments, including Office desktop applications, Office Online in both desktop and mobile browsers, and a growing number of Office tablet and phone add-ins. When you publish your add-ins to the Office Store or to an on-premises add-in catalog, your add-ins will be available to consumers from their Office applications.
To try out some add-ins, you can install the following add-ins from the Office Store.


|**Office product**|**Add-in**|
|:-----|:-----|
|Excel|[Bing Maps](https://store.office.com/bing-maps-WA102957661.aspx?assetid=WA102957661&amp;homapppos=0&amp;homappcat=Data Visualization + BI&amp;homchv=0)|
|Outlook|[Package Tracker](https://store.office.com/package-tracker-WA104162083.aspx?assetid=WA104162083)|
|PowerPoint|[Khan Content from Microsoft](https://store.office.com/khan-content-from-microsoft-WA104320031.aspx?assetid=WA104320031)|
|Word|[Translator](https://store.office.com/translator-WA104124372.aspx?assetid=WA104124372)|
To check out some code, download the [Office Add-ins sample pack](http://code.msdn.microsoft.com/Apps-for-Office-code-d04762b7) for Visual Studio.

## Anatomy of an Office Add-in
<a name="StartBuildingApps_AnatomyofApp"> </a>

The basic components of an Office Add-in are an XML manifest file and the default webpage of your add-in. The manifest defines various settings including the URL of the webpage that implements the add-in's UI and custom logic. When your add-in is ready for your customers, you upload your add-in's manifest to an on-premises add-in catalog or submit it to the Office Store. The webpage (and any .js or other files required by its implementation) needs to be hosted on a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).


**Manifest + webpage = an Office Add-in**

![Manifest plus webpage equals Office Add-in](../images/DK2_AgaveOverview01.png)The manifest specifies settings and capabilities of the add-in, such as the following:



- The URL of the webpage that implements the add-in's UI and programming logic.
    
- The add-in's display name, description, ID, version, and default locale.
    
- How the add-in activates and displays: 
    
      - For add-ins that interact with documents: as a task pane, or in line with document content.
    
  - For add-ins that interact with mail items (messages or appointments): when reading or composing the item.
    
- The permission level and data access requirements for the add-in.
    
For more information, see [Office Add-ins XML manifest](../overview/add-in-manifests.md).


## Development basics
<a name="StartBuildingApps_DevelopmentBasics"> </a>

To create Office Add-ins, you can use any application that can save a file as text. But, you can create an Office Add-in more easily with the Napa Office 365 Development Tools web-based development environment, or in Visual Studio 2015 with its project templates, development environment, and debugging tools. 


### Basic components of an Office Add-in

To create an Office Add-in, at minimum, a developer must create an HTML webpage and a manifest file. The HTML page can be published to any web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md). The manifest file must point to the location of the webpage and be published to any of the following locations: the public Office Store, an internal SharePoint list, or a shared network location.

The most basic Office Add-in consists of a static HTML page that is hosted inside an Office application, but doesn't interact with either the Office document or any other Internet resource. 


**Components of a Hello World Office Add-in**

![Components of a Hello World add-in](../images/DK2_AgaveOverview07.png)


### Creating an Office Add-in with Napa Office 365 Development Tools

Perhaps the quickest way to build an Office Add-in is directly out of a browser. You can do this by using Napa. Napa is web-based development environment that lets you create projects, write code, and run your add-ins all within the browser. There is no need to install any other tools such as Visual Studio. To learn more, see [Create Office Add-ins with Napa with an Office 365 Developer Site](../essentials/create-office-add-ins-with-napa-with-a-developer-site.md). To get started developing right away, see these topics:


- [Create a task pane add-in with Napa Office 365 Development Tools](../essentials/create-a-task-pane-add-in-with-napa.md)
    
- [Create a content add-in for Excel with Napa Office 365 Development Tools](../essentials/create-a-content-add-in-with-napa.md)
    
- [Get Started with Mail add-ins for Office 365](https://dev.outlook.com/MailAppsGettingStarted/GetStarted.aspx)
    
Also, if you begin developing Office Add-ins with Napa, you can develop these projects further in Visual Studio to leverage its more powerful features such as advanced debugging or the ability to use a web project as part of your add-in.


### Creating an Office Add-in with Visual Studio

The most powerful way to build an Office Add-in is to use the  **Add-in for Office** project template in Visual Studio. Visual Studio creates a complete solution that contains all of the files that you need to begin testing your add-in in Office immediately. Visual Studio provides a full range of features to make it easy for you to develop and test Office Add-ins. To learn more, see[Create and debug Office Add-ins in Visual Studio](../essentials/create-and-debug-office-add-ins-in-visual-studio.md). To get started developing right away, see this topic:


- [Create a task pane or content add-in with Visual Studio](../essentials/create-a-task-pane-or-content-add-in-with-visual-studio.md)
    

### Creating an Office Add-in with a text editor

If want to use your favorite text editor to create an Office Add-in, see these topics for information about how to get started:


- [Create a task pane or content add-in for Word or Excel by using a text editor](../essentials/create-a-task-pane-or-content-add-in-for-word-or-excel-by-using-a-text-editor.md)
    
- [Get Started with Mail add-ins for Outlook.com (Preview)](https://dev.outlook.com/MailAppsGettingStarted/GetStarted/outlook-dot-com.aspx)
    

### JavaScript API for Office

The JavaScript API for Office contains objects and members for building add-ins and interacting with Office content and web services.

For more information about the JavaScript API for Office:


- See [Understanding the JavaScript API for Office](../overview/understanding-the-javascript-api-for-office.md) and the[JavaScript API for Office](http://msdn.microsoft.com/library/b27e70c3-d87d-4d27-85e0-103996273298%28Office.15%29.aspx) reference.
    
- Run and edit some JavaScript API for Office code in Excel Online with the [Interactive Office Add-ins API tutorial](http://msdn.microsoft.com/en-us/office/dn449240.aspx)
    
The Word and Excel JavaScript APIs provide host-specific object models that you can use in an Office add-in. You get access to well known objects such as paragraphs and workbooks which makes creating an Office add-in for Word and Excel easier to do. Learn more about these APIs by taking a look at the [Word add-ins](../word/word-add-ins.md) and[Excel add-ins](https://msdn.microsoft.com/EN-US/library/office/mt616485.aspx) overview topics.


## Types of Office Add-ins
<a name="StartBuildingApps_TypesofApps"> </a>

This section provides a quick look at the three types of Office Add-ins: task pane, content, and Outlook. 


### Task pane add-ins

Task pane add-ins work side-by-side with an Office document, and let you supply contextual information and functionality to enhance the document viewing and authoring experience. For example, a task pane add-in can look up and retrieve product information from a web service based on the product name or part number selected in the document.


**Task pane add-in**

![Task Pane add-in](../images/DK2_AgaveOverview04.png)To try out a task pane add-in in Excel 2013, Excel Online, or Word 2013, install the [Wikipedia](https://store.office.com/wikipedia-WA104099688.aspx?assetid=WA104099688) add-in.


### Content add-ins

Content add-ins integrate web-based features as content that shown in line with the body of a document. Content add-ins let you integrate rich, web-based data visualizations, embedded media (such as a YouTube video player or a picture gallery), as well as other external content.


**Content add-in**

![In content add-in](../images/DK2_AgaveOverview05.png)To try out a content add-in in Excel 2013 or Excel Online, install the [Bing Maps](https://store.office.com/bing-maps-WA102957661.aspx?assetid=WA102957661) add-in.


### Outlook add-ins

Outlook add-ins display next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment in a read scenario - the user viewing a received item - or in a compose scenario - the user replying or creating a new item. Outlook add-ins can access contextual information from the item, such as address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification on the various supporting host applications, including Outlook, Outlook for Mac, Outlook Web App and OWA for Devices, to provide a seamless experience on the desktop, web, and tablet and mobile devices.


 **Note**  Outlook add-ins require a minimum version of Exchange 2013 or Exchange Online to host the user's mailbox. POP and IMAP email accounts aren't supported.


**Outlook add-in in a read scenario**

![Contextual add-in](../images/DK2_AgaveOverview06.png)To try out an Outlook add-in in Outlook, Outlook for Mac, or Outlook Web App, install the [Package Tracker](https://store.office.com/package-tracker-WA104162083.aspx?assetid=WA104162083) add-in.


## Office applications that support Office Add-ins
<a name="StartBuildingApps_SupportedApplications"> </a>

Office Add-ins are supported on a growing number of Office host applications running on the desktop, tablets, mobile devices, and in Office Online in the browser. In many cases, this means you can develop a single add-in that runs on different operating systems and Office host applications. And, your customers will have a consistent experience using your add-in across the desktop, their devices, or web browsers.

For task pane add-ins, this means you can develop a single add-in that runs with Excel, PowerPoint, and Word on the Windows desktop, or with Excel Online, PowerPoint Online, Word Online running in a web browser. For Outlook add-ins, this means you can develop a single add-in that runs with Outlook and Outlook for Mac on the desktop, with OWA for Devices on tablet and mobile devices, or with Outlook Web App in a web browser.

This table shows the Office host applications (including desktop, tablet, mobile, and web clients) that can run Office Add-ins, and the types of add-ins supported by each host.


**Supported add-in types**


|**Office application**|**Content add-ins**|**Outlook add-ins**|**Task pane add-ins**|
|:-----|:-----|:-----|:-----|
|Access web apps|
![Check symbol](../images/mod_off15_checkmark.png)

|||
|Excel 2013 or later|
![Check symbol](../images/mod_off15_checkmark.png)

||
![Check symbol](../images/mod_off15_checkmark.png)

|
|Excel Online|
![Check symbol](../images/mod_off15_checkmark.png)

||
![Check symbol](../images/mod_off15_checkmark.png)

|
|Outlook 2013 or later||
![Check symbol](../images/mod_off15_checkmark.png)

||
|Outlook for Mac||
![Check symbol](../images/mod_off15_checkmark.png)

||
|Outlook Web App||
![Check symbol](../images/mod_off15_checkmark.png)

||
|OWA for Devices||
![Check symbol](../images/mod_off15_checkmark.png)

||
|PowerPoint 2013 or later|
![Check symbol](../images/mod_off15_checkmark.png)

||
![Check symbol](../images/mod_off15_checkmark.png)

|
|PowerPoint Online|
![Check symbol](../images/mod_off15_checkmark.png)

||
![Check symbol](../images/mod_off15_checkmark.png)

|
|Project 2013 or later|||
![Check symbol](../images/mod_off15_checkmark.png)

|
|Word 2013 or later|||
![Check symbol](../images/mod_off15_checkmark.png)

|
|Word Online|||
![Check symbol](../images/mod_off15_checkmark.png)

|
For more details, see [Requirements for running Office Add-ins](../overview/requirements-for-running-office-add-ins.md).


## What can an Office Add-in do?
<a name="StartBuildingApps_Capabilities"> </a>

An Office Add-in can do pretty much anything a webpage can do inside the browser, such as the following:


- Provide an interactive UI and custom logic through JavaScript.
    
- Use JavaScript frameworks such as jQuery.
    
- Connect to REST endpoints and web services via HTTP and AJAX.
    
- Run server-side code or logic, if the page is implemented using a server-side scripting language such as ASP or PHP.
    
And, like webpages, Office Add-ins are subject to the same restrictions imposed by browsers, such as the same-origin policy for domain isolation, and security zones. 

In addition to the regular capabilities of a webpage, Office Add-ins can interact with the Office application and an add-in user's content through a JavaScript library that the Office Add-ins infrastructure provides. How your add-ins can interact with Office and content depends on the type of add-in: 


- For task pane and content add-ins, the API lets your add-in read and write to documents, as well as handle key application and user events, such as when the active selection changes. For a summary of the features available to task pane and content add-ins, see [Task pane and content add-ins for Office 2013](../essentials/task-pane-and-content-add-ins.md).
    
- For Outlook add-ins, the API lets your add-in access email message, meeting request, and appointment item properties, and user profile information. The API also provides access to some Exchange Web Services operations. For more information about Outlook add-ins, see [Outlook add-ins](../outlook/outlook-add-ins.md). For a summary of top features of Outlook add-ins, see [Overview of Outlook add-ins architecture and features](../outlook/overview.md).
    

## Understanding the runtime
<a name="StartBuildingApps_Runtime"> </a>

Office Add-ins are secured by an add-in runtime environment, a multiple-tier permissions model, and performance governors. This framework protects the user's experience in the following ways:


- Access to the host application's UI frame is managed.
    
- Only indirect access to the host application's UI thread is allowed.
    
- Modal interactions are not allowed, for example JavaScript alerts aren't allowed.
    
Further, the runtime framework provides the following benefits to ensure that an Office Add-in can't damage an add-in user's environment:


- Isolates the process the add-in runs in.
    
- Doesn't require .dll or .exe replacement or ActiveX components.
    
- Makes add-ins easy to install and uninstall.
    
Also, the runtime framework governs the use of memory, CPU, and network resources by Office Add-ins to ensure that good performance and reliability are maintained. 

For more information about the Office Add-ins privacy and security model, see [Privacy and security for Office Add-ins](87c59a88-10e2-4c88-b6a8-736bd356e5f8.md).


## Publishing basics
<a name="StartBuildingApps_PublishingBasics"> </a>

You can publish Office Add-ins to four distribution end-points:


-  **Office Store**????????This is a public marketplace that Microsoft hosts and regulates on Office.com. In the Office Store, developers around the world can publish and sell their custom Office solutions, and then end users and IT professionals can download them for personal or corporate use. 
    
    When a developer uploads an add-in to the Office Store, Microsoft validates the code. For example, it verifies that the add-in manifest markup is valid and complete. If the code is valid, Microsoft digitally signs the add-in package. The Office Store then takes care of the consumer download experience from discovery to purchase, upgrades, and updates.
    
-  **Office Add-ins catalog on SharePoint**????????For task pane and content add-ins, IT departments can deploy private add-in catalogs to provide the same add-in acquisition experience that the Office Store provides. This new catalog and development platform enables IT departments to use a streamlined method to distribute Office and SharePoint Add-ins to managed users from a central location. 
    
    Add-in catalogs are available to all SharePoint 2013 customers (including Office 365 and SharePoint on-premise). An add-in catalog enables publishing and management of both internally created add-ins as well as add-ins that are available in the Office Store and licensed for corporate use. 
    
-  **Exchange catalog**????????This is a private catalog for Outlook add-ins that is available to users of the Exchange server on which it resides. It enables publishing and management of corporate Outlook add-ins, including internally created add-ins as well as add-ins that are available in the Office Store and licensed for corporate use.
    
-  **Network shared folder add-in catalog**????????IT departments and developers can also deploy task pane and content add-ins to a central network shared folder, where the manifest files will be stored and managed. Users can then acquire add-ins by specifying this shared folder as a trusted catalog, or IT departments can configure this shared folder as a trusted catalog by using a registry setting.
    
For more information, see [Publish your Office Add-in](../publish/publish.md).


## Scenarios
<a name="StartBuildingApps_Scenarios"> </a>

The following scenarios show that Office Add-ins are targeted, quick-hit add-ins that can be used to solve complex, time-consuming problems.

These scenarios suggest ways in which you can, for example, surface line-of-business data and drive adoption of structured business processes in the familiar Office UI across multiple devices. They suggest how you could use an expense-managing add-in that connects Office, SharePoint, and SAP, or create an add-in that combines sales data with maps from the Bing Maps web service to create more effective sales reports. They show how you can unlock the return on your existing investments, such as enterprise resource planning (ERP) and customer relationship management (CRM) applications, by spending less time navigating to and from these applications from an Office client.

Scenarios include:


-  **Translation wizard**????????A Word task pane add-in that automatically translates selected text from the document language to another language selected from a drop-down list.
    
-  **Chart creation**????????An Excel content add-in that builds a chart automatically from selected data.
    
-  **Third-party service integration**????????A Word or Excel task pane add-in that automatically displays the Wikipedia page that corresponds to selected text.
    
-  **Rich mash-ups**????????A Bing map content add-in in Excel that plots the offshore equipment and resource locations for a petroleum company, including getting this information in real time from the company resource-management system.
    
-  **Spec validation**????????A section or paragraph of a design specification for an aircraft component is flagged as outdated, because a Word task pane add-in that communicates with a business system to validate the contents against the latest spec.
    
-  **Kicking off workflows**????????An Outlook add-in can assist creating a message or meeting request based on templates, inserting meeting location details or user's choice of a signature, and attaching related documents.
    
-  **Order details surfaced in context**????????An Outlook add-in that detects a purchase order number or customer number embedded in an email message can present details of the order or customer in the message. This could include an action to take, such as approval.
    

## Additional resources
<a name="StartBuildingApps_AdditionalResources"> </a>


- [Office Add-ins](../overview/office-add-ins.md)
    
- [Design guidelines for Office Add-ins](../design/add-in-design.md)
    
- [Task pane and content add-ins for Office 2013](../essentials/task-pane-and-content-add-ins.md)
    
- [Outlook add-ins](../outlook/outlook-add-ins.md)
    
- [Overview of Outlook add-ins architecture and features](../outlook/overview.md)
    
- [Publish your Office Add-in](../publish/publish.md)
    
- [Office Add-ins API and schema references](../reference/reference.md)
    
